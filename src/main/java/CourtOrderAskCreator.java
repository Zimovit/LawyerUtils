import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;

public abstract class CourtOrderAskCreator {
    static void generate(){

        //I want to let user to return to prev menu, just in case he forgot to get the table
        System.out.println("""
                    Вам необходимо выбрать файл, содержащий таблицу с данными должников.
                    В настоящее время поддерживаются только фойлы Exel.
                    Для выбора файла нажмите "Enter"
                    Для возврата в предыдущее меню введите 0""");

        Scanner scanner = new Scanner(System.in);
        String answer = scanner.nextLine().trim();
        if (answer.equals("0")) return;

        //Lets choose the file that contains the table
        JFileChooser chooser = new JFileChooser();
        //Restrict the choice
        chooser.setAcceptAllFileFilterUsed(false);
        chooser.addChoosableFileFilter(new FileNameExtensionFilter("Microsoft Exel files", "xlsx"));
        int openDialogStatus;
        do {
            openDialogStatus = chooser.showOpenDialog(null);
            if (openDialogStatus != JFileChooser.APPROVE_OPTION){
                System.out.println("Вы не выбрали файл. Хотите вернуться в предыдущее меню? (Д/Н)");
                answer = scanner.nextLine().trim().toLowerCase();
                if (answer.equals("д")) return;
            }
        } while (openDialogStatus != JFileChooser.APPROVE_OPTION);

        //now open the table
        XSSFWorkbook book;
        try {
            book = new XSSFWorkbook(new FileInputStream(chooser.getSelectedFile()));
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Невозможно прочитать файл, проверьте формат файла.\nДля возврата в предыдущее меню нажмите ввод.");
            scanner.nextLine();
            scanner.close();
            return;
        }

        //let's check, the right table has only one sheet
        if (book.getNumberOfSheets() > 1){
            System.out.println("В документе больше одного листа, это неверный формат.\nДля возврата в предыдущее меню нажмите ввод.");
            scanner.nextLine();
            scanner.close();
            return;
        }

        //taking the first sheet
        XSSFSheet sheet = book.getSheetAt(0);

        Iterator<Row> rowIterator = sheet.rowIterator();
        if (!rowIterator.hasNext()){
            System.out.println("Таблица пуста.\nДля возврата в предыдущее меню нажмите ввод.");
            scanner.nextLine();
            return;
        }
        rowIterator.next();   //just skipping the headers

        while (rowIterator.hasNext()){
            XSSFRow row = (XSSFRow) rowIterator.next();
            createDocument(row);
        }


    }

    private static void createDocument(XSSFRow row){
        //collecting data from the row. We have to unify it into an array of strings
        //it is possible that the table is filled wrong, the cells types can be other than the strings
        String[] fieldsOfTheRow = new String[11];
        for (int i = 0; i < fieldsOfTheRow.length; i++){
            XSSFCell cell = row.getCell(i);
            String contentOfTheCell;
            CellType cellType = cell.getCellType();
            switch (cellType){
                case STRING -> contentOfTheCell = cell.getStringCellValue();
                case NUMERIC -> contentOfTheCell = cell.toString();
                default -> contentOfTheCell = "Неверное значение в ячейке";
            }
            fieldsOfTheRow[i] = contentOfTheCell;
        }

        XWPFDocument requestForOrder = new XWPFDocument();
        //fill court name and address
        XWPFParagraph courtNameAndAddress = requestForOrder.createParagraph();
        XWPFRun run = courtNameAndAddress.createRun();
        String text = "Мировому судье судебного участка №1 г. Ельца"+System.lineSeparator()+
                "Елецкого городского судебного района липецкой области.";
        run.setFontSize(14);
        run.setFontFamily("Times New Roman");
        run.setText(text);

        try {
            File file = new File("file.docx");
            requestForOrder.write(new FileOutputStream(file));
            System.out.println(file.getAbsolutePath());
        } catch (IOException e) {
            e.printStackTrace();
        }


    }
}
