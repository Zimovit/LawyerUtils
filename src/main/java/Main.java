import java.util.Scanner;

public class Main {
    public static void main(String[] args) {
        while (true){
            System.out.println("""
                    Программа готова к работе.
                    Выберите действие:
                    1 - Сформировать заявления на судебные приказы по должникам.
                    0 - Выход.""");
            Scanner scanner = new Scanner(System.in);
            String input = scanner.next().trim();
            switch (input) {
                case "1":
                    CourtOrderAskCreator.generate();
                    break;
                case "0":
                    System.exit(0);
                default:
                    System.out.println("Пожалуйста, выберите один из предложенных вариантов.");
            }
        }
    }
}
