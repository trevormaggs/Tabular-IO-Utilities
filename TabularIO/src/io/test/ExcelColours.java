package io.test;

import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelColours
{
    public static void main(String[] args)
    {
        try
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("Default Palettes");
            XSSFCellStyle sampleStyle = null;

            int colourCounter = 0;
        
            for (IndexedColors indexedColor : IndexedColors.values())
            {
                sampleStyle = workbook.createCellStyle();
                sampleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                sampleStyle.setFillForegroundColor(indexedColor.getIndex());

                getCell(sheet, colourCounter, 0).setCellStyle(sampleStyle);
                getCell(sheet, colourCounter, 1).setCellValue(indexedColor.toString());

                colourCounter++;
            }

            FileOutputStream outputStream = new FileOutputStream("POI-Colours.xlsx");
            workbook.write(outputStream);
            workbook.close();
        }
        
        catch (Throwable a_th)
        {
            a_th.printStackTrace();
        }
    }

    private static Cell getCell(XSSFSheet a_sheet, int rowIndex, int cellIndex)
    {
        Row row;
        
        if ((row = a_sheet.getRow(rowIndex)) == null)
        {
            row = a_sheet.createRow(rowIndex);
        }

        Cell cell;
 
        if ((cell = row.getCell(cellIndex)) == null)
        {
            cell = row.createCell(cellIndex);
        }

        return cell;
    }
}