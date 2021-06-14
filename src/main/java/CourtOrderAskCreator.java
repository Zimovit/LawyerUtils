import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Locale;
import java.util.Scanner;

public abstract class CourtOrderAskCreator {

    private static String dirToSaveFiles;

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

        System.out.println("Нажмите ввод и выберите, куда сохранить файлы.");
        scanner.nextLine();

        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        openDialogStatus = chooser.showSaveDialog(null);
        if (openDialogStatus == JFileChooser.APPROVE_OPTION){
            dirToSaveFiles = chooser.getSelectedFile().getAbsolutePath();
        } else {
            System.out.println("Вы не выбрали папку для сохранения, возвращаюсь в начальное меню.");
            return;
        }

        ArrayList<Debtor> listOfDebtors = new ArrayList<>();

        while (rowIterator.hasNext()){
            XSSFRow row = (XSSFRow) rowIterator.next();
            listOfDebtors.add(new Debtor(row));
        }

        //now creating documents iterating the list of debtors

        for (Debtor debtor : listOfDebtors) createDocument(debtor);


    }

    private static void createDocument(Debtor debtor){

        XWPFDocument requestForOrder = new XWPFDocument();
        //fill court name and address
        XWPFParagraph Heading = requestForOrder.createParagraph();
        Heading.setAlignment(ParagraphAlignment.RIGHT);

        XWPFRun run = Heading.createRun();
        run.setFontSize(14);
        run.setFontFamily("Times New Roman");
        run.setText("Мировому судье судебного участка №1 г. Ельца");
        run.addBreak(BreakType.TEXT_WRAPPING);
        run.setText("Елецкого городского судебного района");
        run.addBreak(BreakType.TEXT_WRAPPING);
        run.setText("Липецкой области.");
        run.addBreak(BreakType.TEXT_WRAPPING);
        run.setText("399770, Липецкая обл., г. Елец, ул. Коммунаров, д. 32");

        //now the suitor
        run.addBreak(BreakType.TEXT_WRAPPING);
        run.addBreak(BreakType.TEXT_WRAPPING);
        run.setText("Взыскатель:");
        run.addBreak(BreakType.TEXT_WRAPPING);
        //TODO уточнить
        run.setText("ООО Вентремонт");
        run.addBreak(BreakType.TEXT_WRAPPING);
        run.setText("Полный адрес");

        //debtor
        run.addBreak(BreakType.TEXT_WRAPPING);
        run.addBreak(BreakType.TEXT_WRAPPING);
        run.setText("Должник:");
        run.addBreak(BreakType.TEXT_WRAPPING);
        run.setText(debtor.getName());
        run.addBreak(BreakType.TEXT_WRAPPING);
        run.setText("Дата рождения: " + debtor.getBirthdate().format(DateTimeFormatter.ofPattern("dd MMMM yyyy")) + "г.");
        run.addBreak(BreakType.TEXT_WRAPPING);
        //TODO continue here!
        run.setText("");


        try {
            File file = new File( dirToSaveFiles+"\\"+debtor.getName()+".docx");
            requestForOrder.write(new FileOutputStream(file));
        } catch (IOException e) {
            e.printStackTrace();
        }



    }

}
