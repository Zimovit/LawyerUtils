import java.util.Scanner;

public class Main {
    public static void main(String[] args) {
        while (true){
            System.out.println("""
                    ��������� ������ � ������.
                    �������� ��������:
                    1 - ������������ ��������� �� �������� ������� �� ���������.
                    0 - �����.""");
            Scanner scanner = new Scanner(System.in);
            String input = scanner.next().trim();
            switch (input) {
                case "1":
                    CourtOrderAskCreator.generate();
                    break;
                case "0":
                    System.exit(0);
                default:
                    System.out.println("����������, �������� ���� �� ������������ ���������.");
            }
        }
    }
}
