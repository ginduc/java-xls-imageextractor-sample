package dynobjx;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import java.io.File;
import java.io.FileInputStream;
import java.util.*;

/**
 * Created by ginduc on 4/29/15.
 */
public class Extractor {
    final static int HEADER_ROWNUM = 0;
    final static int ROWNUM = 0;
    final static int TITLE = 1;
    final static int FILENAME = 2;
    final static int IMAGE = 3;

    public static void main(String[] args) throws Exception {
        Extractor extractor = new Extractor();

        File xls = new File(extractor.getClass()
                .getClassLoader().getResource("test.xls").toURI());
        FileInputStream inputStream = null;

        try {
            inputStream = new FileInputStream(xls);
            List<XlsRow> list = extractor.process(inputStream);

            System.out.println("Total elements: " + list.size());

            for (XlsRow xr : list) {
                System.out.println("Row: " + xr);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (inputStream != null) {
                inputStream.close();
            }
            System.exit(0);
        }
    }

    public List<XlsRow> process(FileInputStream inputStream) {
        List<XlsRow> list = new ArrayList<XlsRow>();

        try {
            HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
            HSSFSheet sheet = workbook.getSheetAt(0);

            // Iterate thru spreadsheet rows
            Iterator<Row> rowIterator = sheet.iterator();
            int rowCtr = 0;

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                if (rowCtr > HEADER_ROWNUM) {
                    Cell rowNo = row.getCell(ROWNUM);
                    Cell title = row.getCell(TITLE);
                    Cell filename = row.getCell(FILENAME);

                    if (rowNo != null && title != null) {
                        XlsRow xlsRow = new XlsRow();
                        xlsRow.rowNo = rowNo.getStringCellValue();
                        xlsRow.title = title.getStringCellValue();
                        if (filename != null) {
                            xlsRow.filename = filename.getStringCellValue();
                        }
                        list.add(xlsRow);
                    }
                }

                rowCtr++;
            }

            // Iterate thru images
            HSSFPatriarch patriarch = sheet.getDrawingPatriarch();
            if (patriarch != null) {
                for (HSSFShape shape : patriarch.getChildren()) {
                    if (shape instanceof HSSFPicture) {
                        HSSFPicture picture = (HSSFPicture) shape;
                        Row row = sheet.getRow(picture.getPreferredSize().getRow1());
                        HSSFPictureData pictureData = picture.getPictureData();
                        list.get(row.getRowNum() - 1).imageData =  pictureData.getData();
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return list;
    }


    class XlsRow {
        public String rowNo;
        public String title;
        public String filename;
        public Object imageData;

        @Override
        public String toString() {
            return "XlsRow{" +
                    "rowNo='" + rowNo + '\'' +
                    ", title='" + title + '\'' +
                    ", filename='" + filename + '\'' +
                    ", imageData=" + imageData +
                    '}';
        }
    }
}
