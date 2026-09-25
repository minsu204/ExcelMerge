from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from typing import List, Tuple


class ExcelMerger:
    """엑셀 파일 병합 클래스"""
    
    TITLE_COLUMNS = ['A', 'B', 'C']  # 타이틀 열
    TITLE_ROW = 1  # 타이틀 행

    def __init__(self, file_paths: List[str]):
        """
        Args:
            file_paths: 병합할 엑셀 파일 경로 리스트
        Raises:
            ValueError: 파일 개수가 2개 미만이거나 데이터 위치가 다른 경우
        """
        self.file_paths = file_paths
        self.workbooks = []
        self.validate_files()

    def validate_files(self):
        """파일 검증"""
        # 파일 개수 확인
        if len(self.file_paths) < 2:
            raise ValueError("2개 이상의 파일을 업로드해주세요!")

        # 워크북 로드
        for file_path in self.file_paths:
            try:
                wb = load_workbook(file_path)
                self.workbooks.append(wb)
            except Exception as e:
                raise ValueError(f"파일 읽기 실패: {file_path}\n{str(e)}")

        # 데이터 위치 및 타이틀 검증
        self.validate_data_positions()

    def validate_data_positions(self):
        """
        데이터 위치 및 타이틀 값 검증
        - 모든 파일의 데이터 범위가 같은지 확인
        - A~C 열과 1번 ROW의 타이틀 값이 모두 동일한지 확인
        """
        first_ws = self.workbooks[0].active
        first_max_row = first_ws.max_row
        first_max_col = first_ws.max_column

        # 첫 번째 파일의 타이틀 저장
        first_title_row = {}
        first_title_cols = {}

        for col in self.TITLE_COLUMNS:
            cell = first_ws[f'{col}{self.TITLE_ROW}']
            first_title_row[col] = cell.value

        for row in range(1, first_max_row + 1):
            for col in self.TITLE_COLUMNS:
                cell = first_ws[f'{col}{row}']
                key = f'{col}{row}'
                first_title_cols[key] = cell.value

        # 다른 파일들과 비교
        for i, wb in enumerate(self.workbooks[1:], start=1):
            ws = wb.active
            
            # 데이터 범위 확인
            if ws.max_row != first_max_row or ws.max_column != first_max_col:
                raise ValueError(
                    f"파일 {i+1}의 데이터 범위가 다릅니다!\n"
                    f"파일 1: {first_max_row} x {first_max_col}, "
                    f"파일 {i+1}: {ws.max_row} x {ws.max_column}"
                )

            # 타이틀 행 확인 (A~C 열)
            for col in self.TITLE_COLUMNS:
                cell = ws[f'{col}{self.TITLE_ROW}']
                if cell.value != first_title_row[col]:
                    raise ValueError(
                        f"파일 {i+1}의 {col}{self.TITLE_ROW} 값이 다릅니다!\n"
                        f"파일 1: '{first_title_row[col]}' vs "
                        f"파일 {i+1}: '{cell.value}'"
                    )

            # 타이틀 열 확인 (A~C 열의 모든 행)
            for row in range(1, first_max_row + 1):
                for col in self.TITLE_COLUMNS:
                    cell = ws[f'{col}{row}']
                    key = f'{col}{row}'
                    if cell.value != first_title_cols[key]:
                        raise ValueError(
                            f"파일 {i+1}의 {key} 값이 다릅니다!\n"
                            f"파일 1: '{first_title_cols[key]}' vs "
                            f"파일 {i+1}: '{cell.value}'"
                        )

    def _is_number(self, value):
        """문자열이 숫자인지 확인"""
        if value is None or value == '':
            return False
        if isinstance(value, (int, float)):
            return True
        if isinstance(value, str):
            try:
                float(value)
                return True
            except ValueError:
                return False
        return False

    def _to_number(self, value):
        """값을 숫자로 변환"""
        if value is None or value == '':
            return 0
        if isinstance(value, (int, float)):
            return value
        if isinstance(value, str):
            try:
                if '.' in str(value):
                    return float(value)
                else:
                    return int(value)
            except ValueError:
                return 0
        return 0

    def merge(self):
        """엑셀 파일 병합"""
        result_ws = self.workbooks[0].active

        # 각 파일의 데이터를 순회
        for file_idx, wb in enumerate(self.workbooks[1:], start=1):
            ws = wb.active
            
            # D열부터 마지막 열까지 병합
            for row in range(1, ws.max_row + 1):
                for col_num in range(4, ws.max_column + 1):  # D열(4) 이후부터
                    source_cell = ws.cell(row=row, column=col_num)
                    target_cell = result_ws.cell(row=row, column=col_num)

                    if source_cell.value is not None and source_cell.value != '':
                        # 숫자 문자열인 경우 더함
                        if self._is_number(source_cell.value) and self._is_number(target_cell.value):
                            # 두 값 모두 숫자로 변환 후 더하기
                            target_value = self._to_number(target_cell.value)
                            source_value = self._to_number(source_cell.value)
                            target_cell.value = target_value + source_value
                        # 숫자인데 대상이 숫자가 아닌 경우
                        elif self._is_number(source_cell.value):
                            target_cell.value = self._to_number(source_cell.value)
                        # 그 외의 경우 (주의: 데이터가 덮어씌워질 수 있음)
                        else:
                            if target_cell.value is None or target_cell.value == '':
                                target_cell.value = source_cell.value

        return self.workbooks[0]

    def save(self, output_path: str):
        """병합 결과 저장"""
        try:
            if output_path.lower().endswith('.xls'):
                # XLS 형식으로 저장 (openpyxl -> xlwt 변환)
                self._save_as_xls(output_path)
            else:
                # XLSX 형식으로 저장 (기본)
                self.workbooks[0].save(output_path)
            return True
        except Exception as e:
            raise ValueError(f"저장 실패: {str(e)}")

    def _save_as_xls(self, output_path: str):
        """XLS 형식으로 저장"""
        try:
            import xlwt
        except ImportError:
            raise ValueError("XLS 형식 저장을 위해 xlwt 라이브러리가 필요합니다.\n"
                           "pip install xlwt 를 실행해주세요.")
        
        ws = self.workbooks[0].active
        workbook = xlwt.Workbook()
        worksheet = workbook.add_sheet('Sheet1')
        
        # 모든 셀 데이터 복사
        for row_idx in range(ws.max_row):
            for col_idx in range(ws.max_column):
                cell = ws.cell(row=row_idx + 1, column=col_idx + 1)
                value = cell.value
                worksheet.write(row_idx, col_idx, value)
        
        workbook.save(output_path)
