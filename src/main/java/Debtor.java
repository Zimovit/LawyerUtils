import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;

import java.time.LocalDate;

public class Debtor {

    private String name, birthplace, registrationAddress, premisesAddress, feeFoundation,
            serviceFoundation;

    private double square, fee, debt, poena;

    private LocalDate birthdate, periodStart, periodEnd;

    private FormOfRight premisesFormOfProperty;

    private Stage stage;

    public Debtor(XSSFRow row) throws IllegalArgumentException {
        String cellContent; //is used for all string cell values
        //name
        CellType type = row.getCell(1).getCellType();
        if (type == CellType.STRING){
            this.name = row.getCell(1).getStringCellValue();
        } else {
            throw new IllegalArgumentException("� ������ ��� �������� � �������� �������");
        }

        //birthdate
        try {
            this.birthdate = row.getCell(2).getLocalDateTimeCellValue().toLocalDate();
        } catch (Exception e) {
            throw new IllegalArgumentException("� ������ � ����� �������� �������� ��������");
        }

        //birthplace
        type = row.getCell(3).getCellType();
        if (type == CellType.STRING){
            this.birthplace = row.getCell(3).getStringCellValue();
        } else {
            throw new IllegalArgumentException("� ������ \"����� ��������\" �������� � �������� �������");
        }

        //registrationAddress
        type = row.getCell(4).getCellType();
        if (type == CellType.STRING){
            this.registrationAddress = row.getCell(4).getStringCellValue();
        } else {
            throw new IllegalArgumentException("� ������ \"����� �����������\" �������� � �������� �������");
        }

        //premisesAddress
        type = row.getCell(5).getCellType();
        if (type == CellType.STRING){
            this.premisesAddress = row.getCell(5).getStringCellValue();
        } else {
            throw new IllegalArgumentException("� ������ \"����� ���������\" �������� � �������� �������");
        }

        //premisesFormOfProperty
        type = row.getCell(6).getCellType();
        if (type == CellType.STRING){
            char formOfProperty = row.getCell(6).getStringCellValue().trim().toLowerCase().charAt(0);
            switch (formOfProperty){
                case '�': this.premisesFormOfProperty = FormOfRight.OWNER;
                    break;
                case '�': this.premisesFormOfProperty = FormOfRight.RENTER;
                    break;
                default: throw new IllegalArgumentException("� ������ \"��������� � ���������\" �������� � �������� �������");
            }

        } else {
            throw new IllegalArgumentException("� ������ \"��������� � ���������\" �������� � �������� �������");
        }

        //square
        type = row.getCell(7).getCellType();
        if (type == CellType.NUMERIC){
            this.square = row.getCell(7).getNumericCellValue();
        } else {
            throw new IllegalArgumentException("� ������ \"���������� �������\" �������� � �������� �������");
        }

        //fee
        type = row.getCell(8).getCellType();
        if (type == CellType.NUMERIC){
            this.fee = row.getCell(8).getNumericCellValue();
        } else {
            throw new IllegalArgumentException("� ������ \"�����\" �������� � �������� �������");
        }

        //feeFoundation
        type = row.getCell(9).getCellType();
        if (type == CellType.STRING){
            this.feeFoundation = row.getCell(9).getStringCellValue();
        } else {
            throw new IllegalArgumentException("� ������ \"����������� �������������� ������\" �������� � �������� �������");
        }

        //serviceFoundation
        type = row.getCell(10).getCellType();
        if (type == CellType.STRING){
            this.serviceFoundation = row.getCell(10).getStringCellValue();
        } else {
            throw new IllegalArgumentException("� ������ \"��������� ��� ������������ ����\" �������� � �������� �������");
        }

        //debt
        type = row.getCell(11).getCellType();
        if (type == CellType.NUMERIC){
            this.debt = row.getCell(11).getNumericCellValue();
        } else {
            throw new IllegalArgumentException("� ������ \"����� �������������\" �������� � �������� �������");
        }

        //poena
        type = row.getCell(12).getCellType();
        if (type == CellType.NUMERIC){
            this.poena = row.getCell(12).getNumericCellValue();
        } else {
            throw new IllegalArgumentException("� ������ \"����\" �������� � �������� �������");
        }

        //periodStart
        try {
            this.periodStart = row.getCell(13).getLocalDateTimeCellValue().toLocalDate();
        } catch (Exception e) {
            throw new IllegalArgumentException("� ������ � ����� ������ ���������� ������� �������� ��������");
        }

        //periodEnd
        try {
            this.periodEnd = row.getCell(14).getLocalDateTimeCellValue().toLocalDate();
        } catch (Exception e) {
            throw new IllegalArgumentException("� ������ � ����� ����� ���������� ������� �������� ��������");
        }

        //stage
        type = row.getCell(15).getCellType();
        if (type == CellType.STRING){
            char stageOfCase = row.getCell(15).getStringCellValue().trim().toLowerCase().charAt(0);
            switch (stageOfCase){
                case '�': this.stage = Stage.PRETENSION;
                    break;
                case '�': this.stage = Stage.ORDER;
                    break;
                case '�': this.stage = Stage.SUE;
                    default: throw new IllegalArgumentException("� ������ \"������ ������������\" �������� � �������� �������");
            }

        } else {
            throw new IllegalArgumentException("� ������ \"������ ������������\" �������� � �������� �������");
        }
    }

    public String getName() {
        return name;
    }

    public String getBirthplace() {
        return birthplace;
    }

    public String getRegistrationAddress() {
        return registrationAddress;
    }

    public String getPremisesAddress() {
        return premisesAddress;
    }

    public String getFeeFoundation() {
        return feeFoundation;
    }

    public String getServiceFoundation() {
        return serviceFoundation;
    }

    public double getSquare() {
        return square;
    }

    public double getFee() {
        return fee;
    }

    public double getDebt() {
        return debt;
    }

    public double getPoena() {
        return poena;
    }

    public LocalDate getBirthdate() {
        return birthdate;
    }

    public LocalDate getPeriodStart() {
        return periodStart;
    }

    public LocalDate getPeriodEnd() {
        return periodEnd;
    }

    public FormOfRight getPremisesFormOfProperty() {
        return premisesFormOfProperty;
    }

    public Stage getStage() {
        return stage;
    }

    enum FormOfRight {
        OWNER,
        RENTER
    }

    enum Stage {
        PRETENSION,
        ORDER,
        SUE
    }
}
