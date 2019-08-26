package evgenyt.springdemo;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.servlet.view.document.AbstractXlsView;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Excel file generation
 */

public class ExcelReport extends AbstractXlsView {
    @Override
    protected void buildExcelDocument(Map<String, Object> map,
        Workbook workbook, javax.servlet.http.HttpServletRequest httpServletRequest,
        javax.servlet.http.HttpServletResponse httpServletResponse)
    throws Exception {
        Sheet sheet = workbook.createSheet("Test shit");
        // Generate header
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Col1");
        header.createCell(1).setCellValue("Col2");
        // Get data
        List<String> list = (ArrayList<String>) map.get("testList");
        // Generate rows by data
        for(int idx = 0; idx < list.size(); idx++) {
            Row row = sheet.createRow(idx + 1);
            row.createCell(0).setCellValue(idx);
            row.createCell(1).setCellValue(list.get(idx));
        }
    }
}
