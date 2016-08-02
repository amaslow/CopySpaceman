package copyspaceman;

import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CopySpaceman {

    static String endSrc25 = "_25.jpg";
    static String endSrc26 = "_26.jpg";
    static String endSrc27 = "_27.jpg";
    static String endDst1 = ".1";
    static String endDst2 = ".2";
    static String endDst3 = ".3";
    static String noEAN = "\\NoEAN\\";
    static File dirProductContent = new File("\\\\172.16.55.197\\design\\Smartwares - Product Content\\PRODUCTS\\");
    static File dirDestination = new File("G:\\CM\\Category Management Only\\_S0000_Trade marketing\\Pictures Spaceman\\");
    static String dirArchiveDest = "G:\\CM\\Category Management Only\\_S0000_Trade marketing\\Pictures Spaceman\\Archive+loose pics\\";
    static String excelname = dirDestination + "\\SAP_EAN.xlsx";
    static String excelOutput = dirDestination + "\\UpdateOverview.xlsx";

    public static void main(String[] args) throws IOException {
        FileInputStream fis = null;
        fis = new FileInputStream(excelname);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);

        String[] dirSap = dirProductContent.list(new FilenameFilter() {
            @Override
            public boolean accept(File current, String name) {
                return new File(current, name).isDirectory();
            }
        });
        for (int i = 0; i < dirSap.length; i++) {
            if (dirSap[i].length() == 7) {
                File scrFilePath = new File(dirProductContent + "\\" + dirSap[i] + "\\LR_" + dirSap[i]);
                File src25 = new File(scrFilePath + endSrc25);
                File src26 = new File(scrFilePath + endSrc26);
                File src27 = new File(scrFilePath + endSrc27);
                int counter = 0;

                if (src25.exists()) {
                    counter += 1;
                    File srcNum = src25;
                    String endDst = endDst1;
                    getFromExcel(sheet, srcNum, dirSap[i], endDst, counter);
                }
                if (src26.exists()) {
                    counter += 1;
                    File srcNum = src26;
                    String endDst = endDst2;
                    getFromExcel(sheet, srcNum, dirSap[i], endDst, counter);
                }
                if (src27.exists()) {
                    counter += 1;
                    File srcNum = src27;
                    String endDst = endDst3;
                    getFromExcel(sheet, srcNum, dirSap[i], endDst, counter);
                }
            }
        }
        fis.close();
    }

    private static void getFromExcel(XSSFSheet sheet, File srcNum, String dirSap, String endDst, int counter) throws IOException, FileNotFoundException {
        String sap = dirSap.substring(0, 2) + "." + dirSap.substring(2, 5) + "." + dirSap.substring(5, 7);
        int rownr = findRow(sheet, sap);
        XSSFRow row = sheet.getRow(rownr);
        XSSFCell ean1 = row.getCell(11, Row.RETURN_BLANK_AS_NULL);
        String ean = ean1.getStringCellValue();
        XSSFCell status1 = row.getCell(7, Row.RETURN_BLANK_AS_NULL);
        String status = status1.getStringCellValue();
        if (rownr != 0) {
            if (ean1 != null) {
                String dirEAN = ean.substring(0, 7);
                String fileEAN = ean.substring(7, ean.length());
                File dst = new File(dirDestination + "\\" + dirEAN + "\\" + fileEAN + endDst);
                if (!dst.getParentFile().exists()) {
                    dst.getParentFile().mkdir();
                }
                fileSrcDst(srcNum, dst, ean, sap, status, counter);
            } else {
                File dst = new File(dirDestination + noEAN + dirSap + endDst);
                fileSrcDst(srcNum, dst, ean, sap, status, counter);
            }
        } else {
            File dst = new File(dirDestination + noEAN + dirSap + endDst);
            fileSrcDst(srcNum, dst, ean, sap, status, counter);
        }
    }

    private static void fileSrcDst(File src, File dst, String ean, String sap, String status, int counter) throws IOException, FileNotFoundException {
        DateFormat dateFormat = new SimpleDateFormat("yyyyMMdd");
        File archDest = new File(dirArchiveDest + "\\" + dst.getName());
        if (!dst.exists()) {
            System.out.println("Copy not existing: " + src.getName() + " - into: " + dst);
            copyFile(src, dst);
            if (counter == 1) {
                createLog(ean, sap, status);
            }
        } else {
            if ((new Date(src.lastModified()).after(new Date(new Date().getTime()- (1 * 1000 * 60 * 60 * 24))))) {
                System.out.println("Archive existing: " + dst.getName() + " - into: " + archDest);
                copyFile(dst, archDest);
                System.out.println("... and overwrite: " + src.getName() + " - onto existing: " + dst);
                copyFile(src, dst);
                if (counter == 1) {
                    createLog(ean, sap, status);
                }
            }
        }
    }

    private static void createLog(String ean, String sap, String status) throws IOException, FileNotFoundException {
        FileWriter fw = new FileWriter("H:/Logs/CopySpaceman.log", true);
        BufferedWriter bw = new BufferedWriter(fw);
        DateFormat dateFormater = new SimpleDateFormat("dd-MM-yyyy");
        String modDate = dateFormater.format(new Date());
        System.out.println("Create row in Excel: " + modDate + " - " + ean + " - " + sap + " - " + status);
        bw.newLine();
        bw.write(modDate + " - " + ean + " - " + sap);
        bw.flush();
        bw.close();
        FileInputStream fis = null;
        fis = new FileInputStream(new File(excelOutput));
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);

        int last = (sheet.getLastRowNum() + 1);

        XSSFRow row = sheet.createRow(last);
        XSSFCell modDateCell = row.createCell(0);
        modDateCell.setCellValue(modDate);
        XSSFCell eanCell = row.createCell(1);
        eanCell.setCellValue(ean);
        XSSFCell sapCell = row.createCell(2);
        sapCell.setCellValue(sap);
        XSSFCell statusCell = row.createCell(3);
        statusCell.setCellValue(status);

        fis.close();
        FileOutputStream fos = new FileOutputStream(new File(excelOutput));
        wb.write(fos);
        fos.close();
    }

    private static void copyFile(File src, File dstFile) throws IOException {
        InputStream in = null;
        try {
            in = new FileInputStream(src);
            OutputStream out = new FileOutputStream(dstFile);
            byte[] buf = new byte[1024];
            int len;
            while ((len = in.read(buf)) > 0) {
                out.write(buf, 0, len);
            }
            in.close();
            out.close();
        } catch (FileNotFoundException ex) {
            Logger.getLogger(CopySpaceman.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            in.close();
        }

    }

    private static int findRow(XSSFSheet sheet, String item) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                    if (cell.getRichStringCellValue().getString().trim().equals(item)) {
                        return row.getRowNum();
                    }
                }
            }
        }
        return 0;
    }
}
