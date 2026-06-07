"""
PowerPoint → Word 讲义导出工具

通过 COM 调用 Microsoft Office，将 PPTX 每张幻灯片导出为高清 PNG，
按表格布局自动排版到 Word 文档（左图右备注 / 上图下备注）。

依赖: Python 3.10+, pywin32, Microsoft PowerPoint + Word

运行示例:
    # 默认每页3张，横向，备注在右侧
    python pptx_to_handouts.py "答辩.pptx"

    # 每页2张，纵向，字号12
    python pptx_to_handouts.py "答辩.pptx" --per-page 2 --orientation portrait --font-size 12

    # 每页6张，保留临时图片，前台可见（方便调试）
    python pptx_to_handouts.py "答辩.pptx" --per-page 6 --keep-temp --visible -v

    # 自定义图片分辨率与列宽
    python pptx_to_handouts.py "答辩.pptx" -o "讲义.docx" --slide-width 1920 --notes-width 3.0
"""

from __future__ import annotations

import argparse
import logging
import os
import shutil
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

import pythoncom
import win32com.client
from pywintypes import com_error as COMError


# ==============================================================================
# 配置常量 —— 用户可直接修改此处的默认值，无需触碰业务逻辑
# ==============================================================================

@dataclass
class HandoutConfig:
    """所有可定制参数集中于此。

    用户可直接修改这里的默认值，无需触碰业务逻辑。
    也可以通过 CLI 参数在运行时覆盖。
    """

    # === 布局参数 ===
    per_page: int = 3
    """每页幻灯片数（1 / 2 / 3 / 4 / 6 / 9）"""
    orientation: str = "landscape"
    """页面方向：portrait（纵向）| landscape（横向）。
    当 per_page >= 3 时自动强制为 landscape。"""

    # === 幻灯片导出参数 ===
    slide_width_px: int = 1280
    """导出 PNG 的水平像素数。高度按幻灯片原始宽高比自动计算以避免变形。"""

    # === Word 表格列宽（英寸）===
    slide_col_width_inch: float = 6.2
    """幻灯片图片列宽（英寸）。仅左右布局时生效。"""
    notes_col_width_inch: float = 2.8
    """备注列宽（英寸）。仅左右布局时生效。"""

    # === 备注样式 ===
    font_name: str = "Calibri"
    """备注文本框字体名称。"""
    font_size: int = 10
    """备注文本框字体大小（磅）。"""

    # === 页面边距（英寸）===
    margin_top: float = 0.5
    margin_bottom: float = 0.5
    margin_left: float = 0.8
    margin_right: float = 0.5

    # === 行为控制 ===
    visible: bool = False
    """Word 应用程序是否前台可见（调试时建议设为 True）。"""
    keep_temp: bool = False
    """是否保留临时导出的 PNG 图片（调试时建议设为 True）。"""


# ==============================================================================
# Word COM 常量（不依赖 win32com.client.constants，确保跨环境一致）
# ==============================================================================

class _WdConst:
    """Word 内建枚举常量。"""
    WD_ORIENT_PORTRAIT = 0
    WD_ORIENT_LANDSCAPE = 1
    WD_LINE_STYLE_NONE = 0
    WD_BREAK_PAGE = 7
    WD_COLLAPSE_END = 0
    WD_AUTO_FIT_FIXED = 1       # 固定列宽，不自动调整
    WD_AUTO_FIT_CONTENT = 2
    WD_AUTO_FIT_WINDOW = 3


# ==============================================================================
# 日志配置
# ==============================================================================

logger = logging.getLogger("ppt_handouts")


def _setup_logging(level: int) -> None:
    """配置全局日志格式与级别。"""
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


# ==============================================================================
# PowerPoint 幻灯片提取器
# ==============================================================================

class PPTExtractor:
    """封装 PowerPoint COM 应用，负责导出幻灯片和提取演讲者备注。

    用法:
        extractor = PPTExtractor("slides.pptx", config)
        extractor.export_slide(1, "slide_001.png")
        notes = extractor.get_notes_text(1)
        extractor.close()
    """

    def __init__(self, ppt_path: str, config: HandoutConfig) -> None:
        """打开 PowerPoint 并加载指定演示文稿。

        Args:
            ppt_path: PPTX 文件绝对路径
            config: 讲义全局配置

        Raises:
            RuntimeError: PowerPoint 未安装或 COM 启动失败
        """
        self._ppt_path = ppt_path
        self._config = config
        self._app = None
        self._pres = None

        try:
            self._app = win32com.client.Dispatch("PowerPoint.Application")
        except Exception as exc:
            raise RuntimeError(
                f"无法启动 PowerPoint COM 应用。请确认已安装 Microsoft PowerPoint。\n"
                f"原始错误: {exc}"
            ) from exc

        # 某些环境下不可设置（如 PowerPoint 已在运行），安全忽略
        try:
            self._app.Visible = False
        except Exception:
            pass
        try:
            self._app.DisplayAlerts = 0  # ppAlertsNone
        except Exception:
            pass

        ppt_abs = os.path.abspath(ppt_path)
        try:
            self._pres = self._app.Presentations.Open(ppt_abs, WithWindow=False)
        except Exception as exc:
            self._safe_quit()
            raise RuntimeError(f"无法打开演示文稿: {ppt_abs}\n原始错误: {exc}") from exc

        logger.info("已打开演示文稿，共 %d 张幻灯片", self.slide_count)

    # ---- 属性 ---------------------------------------------------------------

    @property
    def slide_count(self) -> int:
        """幻灯片总数。"""
        return len(self._pres.Slides)

    # ---- 导出幻灯片 ---------------------------------------------------------

    def export_slide(self, index: int, png_path: str) -> None:
        """将指定幻灯片导出为 PNG 图片。

        Args:
            index: 幻灯片序号（1-based）
            png_path: 输出 PNG 文件路径
        """
        slide = self._pres.Slides(index)
        w, h = self._compute_export_size()
        slide.Export(png_path, "PNG", w, h)

    def _compute_export_size(self) -> tuple[int, int]:
        """根据幻灯片原始宽高比计算导出像素尺寸，避免图片拉伸变形。"""
        ps = self._pres.PageSetup
        slide_w, slide_h = ps.SlideWidth, ps.SlideHeight
        if slide_w <= 0:
            return self._config.slide_width_px, self._config.slide_width_px * 9 // 16
        aspect = slide_h / slide_w
        w = self._config.slide_width_px
        h = int(w * aspect)
        return w, h

    # ---- 读取备注 ----------------------------------------------------------

    def get_notes_text(self, index: int) -> str:
        """读取指定幻灯片的演讲者备注文本。

        Args:
            index: 幻灯片序号（1-based）

        Returns:
            备注文本。若无备注或读取失败则返回空字符串。
        """
        try:
            slide = self._pres.Slides(index)
            notes_page = slide.NotesPage
            # Placeholders(2) 是备注正文占位符（Placeholders(1) 为备注页标题）
            text = notes_page.Shapes.Placeholders(2).TextFrame.TextRange.Text
            return text.strip() if text else ""
        except Exception:
            return ""

    # ---- 资源释放 ----------------------------------------------------------

    def close(self) -> None:
        """关闭 PowerPoint 演示文稿并退出应用，释放 COM 资源。"""
        if self._pres is not None:
            try:
                self._pres.Close()
            except Exception:
                pass
            self._pres = None
        self._safe_quit()

    def _safe_quit(self) -> None:
        """安全退出 PowerPoint 应用程序。"""
        if self._app is not None:
            try:
                self._app.Quit()
            except Exception:
                pass
            self._app = None

    def __del__(self) -> None:
        self.close()


# ==============================================================================
# Word 讲义组装器
# ==============================================================================

class WordAssembler:
    """封装 Word COM 应用，负责创建表格布局讲义文档。

    用法:
        assembler = WordAssembler(config)
        assembler.set_page_orientation("landscape")
        table = assembler.create_table(3, 2)
        assembler.insert_image(table.Cell(1, 1), "slide_001.png", 300)
        assembler.insert_notes(table.Cell(1, 2), "备注文字")
        assembler.save_and_quit("output.docx")
    """

    def __init__(self, config: HandoutConfig) -> None:
        """启动 Word 并创建空白文档。

        Args:
            config: 讲义全局配置
        """
        self._config = config
        self._app = None
        self._doc = None

        try:
            self._app = win32com.client.Dispatch("Word.Application")
        except Exception as exc:
            raise RuntimeError(
                f"无法启动 Word COM 应用。请确认已安装 Microsoft Word。\n"
                f"原始错误: {exc}"
            ) from exc

        self._app.Visible = config.visible
        try:
            self._app.DisplayAlerts = 0  # wdAlertsNone
        except Exception:
            pass

        self._doc = self._app.Documents.Add()
        logger.info("已创建空白 Word 文档（界面%s可见）", "" if config.visible else "不")

    # ---- 页面设置 ----------------------------------------------------------

    def set_page_orientation(self, orientation: str) -> None:
        """设置页面方向。

        Args:
            orientation: "portrait" 或 "landscape"
        """
        if orientation == "landscape":
            self._doc.PageSetup.Orientation = _WdConst.WD_ORIENT_LANDSCAPE
        else:
            self._doc.PageSetup.Orientation = _WdConst.WD_ORIENT_PORTRAIT

    # ---- 单位转换 ----------------------------------------------------------

    @staticmethod
    def inches_to_points(inches: float) -> float:
        """英寸转磅（1 inch = 72 points）。

        使用纯数学计算，避免 COM 动态绑定的 InchesToPoints 不可用。
        """
        return inches * 72.0

    def set_page_margins(self) -> None:
        """应用配置中的页面边距。"""
        cfg = self._config
        ps = self._doc.PageSetup
        pt = self.inches_to_points
        ps.TopMargin = pt(cfg.margin_top)
        ps.BottomMargin = pt(cfg.margin_bottom)
        ps.LeftMargin = pt(cfg.margin_left)
        ps.RightMargin = pt(cfg.margin_right)

    # ---- 页面尺寸计算 -------------------------------------------------------

    def get_page_usable_size_pts(self) -> tuple[float, float]:
        """返回页面可用区域宽高（磅），已扣除页边距。"""
        ps = self._doc.PageSetup
        usable_w = ps.PageWidth - ps.LeftMargin - ps.RightMargin
        usable_h = ps.PageHeight - ps.TopMargin - ps.BottomMargin
        return usable_w, usable_h

    # ---- 表格 ------------------------------------------------------------

    def create_table_at_end(self, rows: int, cols: int):
        """在文档末尾创建表格并返回 Table 对象。

        表格默认 AutoFit 行为设为 Fixed，避免列宽被自动调整。

        Args:
            rows: 行数
            cols: 列数

        Returns:
            Word Table COM 对象
        """
        rng = self._doc.Range()
        rng.Collapse(_WdConst.WD_COLLAPSE_END)
        table = self._doc.Tables.Add(rng, rows, cols)
        # 锁定列宽为固定模式，防止 Word 自动伸缩
        try:
            table.AutoFitBehavior(_WdConst.WD_AUTO_FIT_FIXED)
        except Exception:
            pass
        return table

    def set_column_widths(self, table, widths_pts: list[float]) -> None:
        """按列表顺序设置各列宽度（磅）。

        Args:
            table: Word Table COM 对象
            widths_pts: 各列宽度列表
        """
        for i, w in enumerate(widths_pts, start=1):
            try:
                table.Columns(i).Width = w
            except Exception:
                pass

    def set_cell_padding(self, table, top: float = 2.0, bottom: float = 2.0,
                         left: float = 4.0, right: float = 4.0) -> None:
        """统一设置表格所有单元格内边距（磅）。

        Args:
            table: Word Table COM 对象
            top / bottom / left / right: 内边距（磅）
        """
        try:
            table.TopPadding = top
            table.BottomPadding = bottom
            table.LeftPadding = left
            table.RightPadding = right
        except Exception:
            pass

    def remove_borders(self, table) -> None:
        """移除表格所有可见边框线。"""
        try:
            table.Borders.InsideLineStyle = _WdConst.WD_LINE_STYLE_NONE
            table.Borders.OutsideLineStyle = _WdConst.WD_LINE_STYLE_NONE
        except Exception:
            pass

    # ---- 内容填充 ----------------------------------------------------------

    def insert_image(self, cell, image_path: str,
                     max_width_pts: float, max_height_pts: float = 99999) -> None:
        """在表格单元格中插入嵌入式图片，约束宽高不超出限制，等比缩放。

        Args:
            cell: Word Cell COM 对象
            image_path: PNG 文件路径
            max_width_pts: 最大宽度限制（磅）
            max_height_pts: 最大高度限制（磅），默认无限制
        """
        if not os.path.exists(image_path):
            logger.warning("图片不存在，跳过: %s", image_path)
            return

        shape = cell.Range.InlineShapes.AddPicture(
            FileName=os.path.abspath(image_path),
            LinkToFile=False,
            SaveWithDocument=True,
        )

        orig_w = shape.Width
        orig_h = shape.Height
        if orig_w <= 0:
            return

        ratio = orig_h / orig_w if orig_w > 0 else 1.0

        # 同时约束宽和高，选择更严格的缩放比
        scale_w = max_width_pts / orig_w if orig_w > max_width_pts else 1.0
        scale_h = max_height_pts / orig_h if orig_h > max_height_pts else 1.0
        scale = min(scale_w, scale_h)

        if scale < 1.0:
            shape.Width = orig_w * scale
            shape.Height = orig_h * scale

    def insert_notes(self, cell, text: str) -> None:
        """在表格单元格中插入备注文本并应用字体样式。

        Args:
            cell: Word Cell COM 对象
            text: 备注文本
        """
        if not text:
            return

        cfg = self._config
        cell_range = cell.Range
        cell_range.Text = text

        # 对单元格内所有段落统一应用字体
        for paragraph in cell_range.Paragraphs:
            paragraph.Range.Font.Name = cfg.font_name
            paragraph.Range.Font.Size = cfg.font_size
            paragraph.Format.SpaceAfter = 4

    # ---- 分页 ---------------------------------------------------------------

    def insert_page_break(self) -> None:
        """在文档末尾插入分页符。"""
        rng = self._doc.Range()
        rng.Collapse(_WdConst.WD_COLLAPSE_END)
        rng.InsertBreak(_WdConst.WD_BREAK_PAGE)

    # ---- 保存与退出 ---------------------------------------------------------

    def save_and_quit(self, output_path: str) -> None:
        """保存文档到指定路径并退出 Word 应用。

        Args:
            output_path: 输出 DOCX 文件绝对路径
        """
        try:
            self._doc.SaveAs(os.path.abspath(output_path))
            logger.info("讲义已保存至: %s", output_path)
        finally:
            self._close()

    def _close(self) -> None:
        """关闭 Word 文档和应用程序实例。"""
        if self._doc is not None:
            try:
                self._doc.Close(SaveChanges=False)
            except Exception:
                pass
            self._doc = None
        if self._app is not None:
            try:
                self._app.Quit()
            except Exception:
                pass
            self._app = None

    def __del__(self) -> None:
        self._close()


# ==============================================================================
# 核心编排逻辑
# ==============================================================================

def build_handouts(
    ppt_path: str,
    docx_path: str,
    config: HandoutConfig,
) -> None:
    """主编排函数：遍历幻灯片、导出 PNG、在 Word 中排版生成讲义。

    进程守护：
        无论正常结束还是异常中断，finally 块保证 COM 资源释放和临时文件清理。

    Args:
        ppt_path: 输入 PPTX 文件路径
        docx_path: 输出 DOCX 文件路径
        config: 讲义配置
    """
    temp_dir: Optional[str] = None
    ppt_extractor: Optional[PPTExtractor] = None
    word_assembler: Optional[WordAssembler] = None

    try:
        # ================================================================
        # 阶段 1：导出所有幻灯片为 PNG
        # ================================================================
        ppt_extractor = PPTExtractor(ppt_path, config)
        total_slides = ppt_extractor.slide_count

        temp_dir = tempfile.mkdtemp(prefix="pptx_handouts_")
        logger.info("临时图片目录: %s", temp_dir)

        png_paths: List[str] = []
        notes_list: List[str] = []

        retry_delay = 0.5  # 幻灯片间延迟（秒），防止 COM 消息队列拥塞
        max_retries = 3    # 单张幻灯片最大重试次数

        for i in range(1, total_slides + 1):
            # 每张幻灯片间短暂休眠，给 COM 消息泵处理时间
            if i > 1:
                time.sleep(retry_delay)
                pythoncom.PumpWaitingMessages()

            png_path = os.path.join(temp_dir, f"slide_{i:03d}.png")
            notes_list.append(ppt_extractor.get_notes_text(i))

            exported = False
            for attempt in range(1, max_retries + 1):
                try:
                    logger.info("正在导出幻灯片 %d / %d (第 %d 次)...",
                                i, total_slides, attempt)
                    ppt_extractor.export_slide(i, png_path)
                    png_paths.append(png_path)
                    exported = True
                    break
                except Exception as exc:
                    logger.warning("幻灯片 %d 第 %d 次导出失败: %s", i, attempt, exc)
                    if attempt < max_retries:
                        # 指数退避：0.5s → 1.0s → 2.0s
                        wait = retry_delay * (2 ** (attempt - 1))
                        logger.info("将在 %.1f 秒后重试...", wait)
                        time.sleep(wait)
                        pythoncom.PumpWaitingMessages()
                    else:
                        logger.warning("幻灯片 %d 最终导出失败，将跳过", i)
                        png_paths.append("")  # 占位，保持索引对齐

        # 关闭 PowerPoint，Word 阶段不再需要
        ppt_extractor.close()
        ppt_extractor = None

        valid_slides = sum(1 for p in png_paths if p)
        if valid_slides == 0:
            raise RuntimeError("所有幻灯片导出均失败，无法生成讲义。")

        logger.info("幻灯片导出完毕: 成功 %d / 共 %d 张", valid_slides, total_slides)

        # ================================================================
        # 阶段 2：确定布局方案
        # ================================================================
        per_page = config.per_page
        effective_orientation = config.orientation

        # per_page=1 + portrait → 上下布局（1 列 2 行）
        is_top_bottom = (per_page == 1 and effective_orientation == "portrait")
        layout_name = "纵向-上下" if is_top_bottom else f"横向-左右({per_page}张/页)"
        logger.info("布局模式: %s | 页面方向: %s", layout_name, effective_orientation)

        # ================================================================
        # 阶段 3：初始化 Word 文档
        # ================================================================
        word_assembler = WordAssembler(config)
        word_assembler.set_page_orientation(effective_orientation)
        word_assembler.set_page_margins()

        usable_w, usable_h = word_assembler.get_page_usable_size_pts()

        # ---- 计算列宽和行高约束 ----
        if is_top_bottom:
            col_widths = [usable_w]
            img_max_w = usable_w * 0.95
            # per_page=1 纵向：图片占页面可用高度的 70%，留 30% 给备注
            img_max_h = usable_h * 0.70
        else:
            img_col_w = word_assembler.inches_to_points(config.slide_col_width_inch)
            notes_col_w_raw = word_assembler.inches_to_points(config.notes_col_width_inch)
            total = img_col_w + notes_col_w_raw
            if total > usable_w * 1.01:
                scale = usable_w / total
                img_col_w *= scale
                notes_col_w_raw *= scale
            col_widths = [img_col_w, notes_col_w_raw]
            img_max_w = img_col_w * 0.95
            # 每张幻灯片占页面高度的 1/per_page，留 15% 给行间距/备注空间
            img_max_h = (usable_h / per_page) * 0.85

        # ================================================================
        # 阶段 4：逐页生成表格
        # ================================================================
        slide_idx = 0
        all_count = len(png_paths)
        page_num = 0

        while slide_idx < all_count:
            page_num += 1
            remaining = all_count - slide_idx
            rows_this_page = min(per_page, remaining)

            logger.info("正在生成第 %d 页讲义（幻灯片 %d-%d）...",
                        page_num, slide_idx + 1, slide_idx + rows_this_page)

            if is_top_bottom:
                # ---- 纵向上下布局: 图片在上 / 备注在下，1 列表格 ----
                table = word_assembler.create_table_at_end(2, 1)
                word_assembler.set_column_widths(table, col_widths)
                word_assembler.set_cell_padding(table)
                word_assembler.remove_borders(table)

                png = png_paths[slide_idx]
                notes = notes_list[slide_idx]

                if png:
                    word_assembler.insert_image(table.Cell(1, 1), png, img_max_w, img_max_h)
                word_assembler.insert_notes(table.Cell(2, 1), notes)
                slide_idx += 1
            else:
                # ---- 横向左右布局: 图片在左 / 备注在右，N 行 2 列表格 ----
                table = word_assembler.create_table_at_end(rows_this_page, 2)
                word_assembler.set_column_widths(table, col_widths)
                word_assembler.set_cell_padding(table)
                word_assembler.remove_borders(table)

                for row in range(1, rows_this_page + 1):
                    png = png_paths[slide_idx]
                    notes = notes_list[slide_idx]

                    if png:
                        word_assembler.insert_image(table.Cell(row, 1), png, img_max_w, img_max_h)
                    word_assembler.insert_notes(table.Cell(row, 2), notes)
                    slide_idx += 1

        # ================================================================
        # 阶段 5：保存并退出
        # ================================================================
        word_assembler.save_and_quit(docx_path)
        word_assembler = None  # 防止 finally 重复关闭

        logger.info("完成! 共生成 %d 页讲义，覆盖 %d 张幻灯片 → %s",
                    page_num, all_count, docx_path)

    finally:
        # ---- 进程守护：无论成功与失败都释放 COM 和临时文件 ----
        if ppt_extractor is not None:
            ppt_extractor.close()
        if word_assembler is not None:
            word_assembler._close()

        if temp_dir is not None:
            if config.keep_temp:
                logger.info("临时图片已保留在: %s", temp_dir)
            else:
                try:
                    shutil.rmtree(temp_dir, ignore_errors=True)
                    logger.debug("已清理临时目录: %s", temp_dir)
                except Exception:
                    pass


# ==============================================================================
# CLI 入口
# ==============================================================================

def main() -> None:
    """命令行入口函数。解析参数、校验输入、调用讲义编排。"""

    parser = argparse.ArgumentParser(
        description="将 PowerPoint 演示文稿导出为 Word 讲义文档（表格布局，左图右备注）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python pptx_to_handouts.py "答辩.pptx"
  python pptx_to_handouts.py "答辩.pptx" -o "讲义.docx" --per-page 2 --orientation portrait
  python pptx_to_handouts.py "答辩.pptx" --per-page 6 --font-size 12 --visible -v
  python pptx_to_handouts.py "答辩.pptx" --slide-width 1920 --notes-width 3.0 --keep-temp
        """,
    )

    parser.add_argument(
        "input",
        help="输入 PPTX 文件路径（必填）",
    )
    parser.add_argument(
        "--output", "-o",
        default=None,
        help="输出 DOCX 文件路径。默认与输入同目录，文件名添加 _handouts 后缀。",
    )
    parser.add_argument(
        "--per-page",
        type=int,
        default=HandoutConfig.per_page,
        choices=[1, 2, 3, 4, 6, 9],
        help=f"每页幻灯片数（默认: {HandoutConfig.per_page}）",
    )
    parser.add_argument(
        "--orientation",
        default=HandoutConfig.orientation,
        choices=["portrait", "landscape"],
        help=f"页面方向（默认: {HandoutConfig.orientation}）。per_page >= 3 时自动强制 landscape。",
    )
    parser.add_argument(
        "--slide-width",
        type=int,
        default=HandoutConfig.slide_width_px,
        help=f"导出 PNG 宽度，像素（默认: {HandoutConfig.slide_width_px}）",
    )
    parser.add_argument(
        "--slide-height",
        type=int,
        default=0,
        help="导出 PNG 高度（已弃用，高度按宽高比自动计算以保持原始比例）",
    )
    parser.add_argument(
        "--notes-width",
        type=float,
        default=HandoutConfig.notes_col_width_inch,
        help=f"备注列宽度，英寸（默认: {HandoutConfig.notes_col_width_inch}）",
    )
    parser.add_argument(
        "--slide-width-inch",
        type=float,
        default=HandoutConfig.slide_col_width_inch,
        help=f"幻灯片图片列宽度，英寸（默认: {HandoutConfig.slide_col_width_inch}）",
    )
    parser.add_argument(
        "--font-name",
        default=HandoutConfig.font_name,
        help=f"备注字体名称（默认: {HandoutConfig.font_name}）",
    )
    parser.add_argument(
        "--font-size",
        type=int,
        default=HandoutConfig.font_size,
        help=f"备注字号，磅（默认: {HandoutConfig.font_size}）",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        default=HandoutConfig.visible,
        help="Word 应用程序前台可见（调试时使用）",
    )
    parser.add_argument(
        "--keep-temp",
        action="store_true",
        default=HandoutConfig.keep_temp,
        help="保留临时导出的 PNG 文件（调试时使用）",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        default=False,
        help="打印详细进度日志",
    )

    args = parser.parse_args()

    # ---- 日志 ----
    _setup_logging(logging.DEBUG if args.verbose else logging.INFO)

    # ---- 输入校验 ----
    input_path = Path(args.input).resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在: {input_path}")
    if input_path.suffix.lower() not in (".pptx", ".ppt"):
        raise ValueError(f"输入文件不是 PowerPoint 格式 (.pptx / .ppt): {input_path}")

    # ---- 输出路径 ----
    if args.output:
        output_path = Path(args.output).resolve()
    else:
        output_path = input_path.parent / f"{input_path.stem}_handouts.docx"

    output_dir = output_path.parent
    if not output_dir.exists():
        raise FileNotFoundError(f"输出目录不存在: {output_dir}")

    # ---- 组装配置 ----
    config = HandoutConfig(
        per_page=args.per_page,
        orientation=args.orientation,
        slide_width_px=args.slide_width,
        slide_col_width_inch=args.slide_width_inch,
        notes_col_width_inch=args.notes_width,
        font_name=args.font_name,
        font_size=args.font_size,
        visible=args.visible,
        keep_temp=args.keep_temp,
    )

    logger.info("输入:  %s", input_path)
    logger.info("输出:  %s", output_path)
    logger.info("参数:  per_page=%d  orientation=%s  font=%s(%dpt)",
                config.per_page, config.orientation, config.font_name, config.font_size)

    # ---- 执行 ----
    try:
        build_handouts(str(input_path), str(output_path), config)
    except COMError as exc:
        logger.error("COM 接口调用失败")
        raise RuntimeError(
            "无法连接 Microsoft Office COM 接口。\n"
            "请确保已安装 Microsoft PowerPoint 和 Microsoft Word，且 COM 注册状态正常。\n"
            "可尝试以管理员身份运行:\n"
            '  & ".venv\\Scripts\\python.exe" "Scripts\\pywin32_postinstall.py" -install\n'
            f"原始 COM 错误: {exc}"
        ) from exc


if __name__ == "__main__":
    main()
