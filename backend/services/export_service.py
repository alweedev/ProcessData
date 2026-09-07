import io

import pandas as pd


class ExportService:
    @staticmethod
    def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Dados") -> io.BytesIO:
        output = io.BytesIO()
        try:
            import openpyxl  # noqa: F401
            from openpyxl.styles import Alignment, Font, PatternFill
            from openpyxl.utils import get_column_letter

            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]

                header_fill = PatternFill(
                    start_color="FFDCE6F1",
                    end_color="FFDCE6F1",
                    fill_type="solid",
                )
                for cell in list(ws[1]):
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.fill = header_fill

                for idx, col in enumerate(df.columns, 1):
                    series = df[col].astype(str).fillna("")
                    max_len = max(series.map(len).max(), len(str(col))) + 2
                    max_len = min(max_len, 60)
                    ws.column_dimensions[get_column_letter(idx)].width = max_len

                ws.freeze_panes = "A2"
                try:
                    ws.auto_filter.ref = ws.dimensions
                except Exception:
                    pass

            output.seek(0)
        except Exception:
            output = io.BytesIO()
            df.to_excel(output, index=False)
            output.seek(0)
        return output
