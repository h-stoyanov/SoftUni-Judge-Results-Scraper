import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.TreeMap;

public class SheetUtil {
    private static String fileWithPath = "src/main/resources/students.xlsx";

    /**
     * This method is returning all courses with their column indexes.
     *
     * @return Map with Key course name and Value column index.
     * @throws IOException            if file is not found.
     * @throws InvalidFormatException if file is not in valid format.
     */
    static TreeMap<String, Integer> getAllLectures() throws IOException, InvalidFormatException {
        TreeMap<String, Integer> courses = new TreeMap<>();
        InputStream inputStream = new FileInputStream(fileWithPath);
        Workbook workbook = WorkbookFactory.create(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        Row head = sheet.getRow(0);
        for (Cell cell : head) {
            if (cell.getColumnIndex() > 1)
                courses.put(cell.getStringCellValue(), cell.getColumnIndex());
        }
        inputStream.close();
        return courses;
    }

    /**
     * This method is returning array of integers with Lecture info.
     *
     * @param columnId is the column id of Lecture in the spreadsheet.
     * @return array of integers on index 0 is the id in judge, index 1 is the max points for the lecture
     * @throws IOException            if file is not found.
     * @throws InvalidFormatException if file is not in valid format.
     */
    static int[] getLectureInfoByColumnId(int columnId) throws IOException, InvalidFormatException {
        int[] info = new int[3];
        InputStream inputStream = new FileInputStream(fileWithPath);
        Workbook workbook = WorkbookFactory.create(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        info[0] = columnId;
        info[1] = (int) sheet.getRow(1).getCell(columnId).getNumericCellValue();
        info[2] = (int) sheet.getRow(2).getCell(columnId).getNumericCellValue();
        return info;
    }

    /**
     * This method is returning all students with their row indexes.
     *
     * @return Map with Key student username and Value row index.
     * @throws IOException            if file is not found.
     * @throws InvalidFormatException if file is not in valid format.
     */
    static Map<String, Integer> getAllStudents() throws IOException, InvalidFormatException {
        Map<String, Integer> students = new HashMap<>();
        InputStream inputStream = new FileInputStream(fileWithPath);
        Workbook workbook = WorkbookFactory.create(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        for (Row row : sheet) {
            if (row.getRowNum() > 2) {
                Cell cell = row.getCell(1);
                if (cell != null)
                    students.put(cell.getStringCellValue(), row.getRowNum());
            }
        }
        inputStream.close();
        return students;
    }

    /**
     * This method will be used to set value with student result to specific cell.
     *
     * @param rowId is the id of the row which the user corresponds to.
     * @param colId is the id of the column which the lecture corresponds to.
     * @param value is the value which has to be written to the cell
     * @throws IOException            if file is not found.
     * @throws InvalidFormatException if file is not in valid format.
     */
    static void writeResultToCell(int rowId, int colId, double value) throws IOException, InvalidFormatException {
        InputStream inputStream = new FileInputStream(fileWithPath);
        Workbook workbook = WorkbookFactory.create(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(rowId);
        Cell cell = row.getCell(colId);
        if (cell == null)
            cell = row.createCell(colId);
        cell.setCellValue(value);
        inputStream.close();
        FileOutputStream fileOut = new FileOutputStream(fileWithPath);
        workbook.write(fileOut);
        fileOut.close();
    }
}
