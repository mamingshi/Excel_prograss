"""读取 Excel 工作簿；所有操作均为只读，不保存原始文件。"""

from dataclasses import dataclass
import hashlib
from pathlib import Path
from zipfile import BadZipFile

import pandas as pd


def file_sha256(path: Path) -> str:
    """只读计算文件身份，防止导出到另一版本的源工作簿。"""
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


class ExcelLoadError(Exception):
    """包含中文提示及原始错误详情的读取异常。"""

    def __init__(self, message: str, detail: str = "") -> None:
        super().__init__(message)
        self.detail = detail


@dataclass
class ExcelPreview:
    """一次读取的预览结果，供界面和后续业务使用。"""

    file_path: Path
    sheet_names: tuple[str, ...]
    sheet_name: str
    dataframe: pd.DataFrame
    source_digest: str | None = None


class ExcelFileLoader:
    """负责文件校验、工作表枚举和单个工作表读取。"""

    def load(self, file_path: str | Path, sheet_name: str | None = None) -> ExcelPreview:
        path = Path(file_path)
        if path.suffix.lower() != ".xlsx":
            raise ExcelLoadError("当前版本仅支持 .xlsx 格式，请选择 Excel 工作簿。")

        try:
            if not path.is_file():
                raise ExcelLoadError("文件不存在或已被移动，请重新选择文件。")

            source_digest = file_sha256(path)

            # 上下文管理器及时关闭文件，切换文件后不会遗留文件句柄。
            with pd.ExcelFile(path, engine="openpyxl") as workbook:
                sheet_names = tuple(workbook.sheet_names)
                if not sheet_names:
                    raise ExcelLoadError("该 Excel 文件没有可读取的工作表。")

                selected_sheet = sheet_name if sheet_name is not None else sheet_names[0]
                if selected_sheet not in sheet_names:
                    raise ExcelLoadError("所选工作表已不存在，请重新打开文件。")

                # 第一行作为字段名；保留文本类型及 NA 等原始文本。
                dataframe = self.read_sheet(workbook, selected_sheet)

            if file_sha256(path) != source_digest:
                raise ExcelLoadError("文件在读取期间发生变化，请关闭其他编辑程序后重新打开。")
            return ExcelPreview(path.resolve(), sheet_names, selected_sheet, dataframe, source_digest)
        except ExcelLoadError:
            raise
        except PermissionError as error:
            raise ExcelLoadError(
                "无法访问此文件，请检查文件访问权限，或关闭占用文件的程序后重试。",
                str(error),
            ) from error
        except FileNotFoundError as error:
            raise ExcelLoadError("文件不存在或已被移动，请重新选择文件。", str(error)) from error
        except BadZipFile as error:
            raise ExcelLoadError(
                "文件不是有效的 .xlsx 工作簿，可能已损坏或受到密码保护。",
                str(error),
            ) from error
        except Exception as error:
            raise ExcelLoadError(
                "读取 Excel 失败，请确认文件有效、未损坏且未设置打开密码。",
                str(error),
            ) from error

    @staticmethod
    def read_sheet(workbook: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
        """统一读取约定：首行表头、完整列顺序、文本不作数字转换。"""
        return workbook.parse(sheet_name=sheet_name, header=0, dtype=object, keep_default_na=False)
