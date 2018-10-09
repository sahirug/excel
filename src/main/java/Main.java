import java.io.File;
import java.io.IOException;
import java.text.MessageFormat;

public class Main {

    private final static String EXCEL_PATH = "C:\\Users\\sahir\\Documents\\Macros\\MacroWS.xlsm";

    private final static String MACRO_NAME = "NewMacro2";

    private final static String FILE_NAME = "MacroWS.xlsm";

    public static void main(String[] args) {
        String path = new File(Main.class.getClassLoader().getResource("myvbs.vbs").getFile()).getAbsolutePath();
        String cmd = MessageFormat.format("wscript {0} {1} {2} {3}", path, EXCEL_PATH, FILE_NAME, MACRO_NAME);

//        String cmd = "wscript " + path;
        test2(cmd);
    }

    private static void test1() {
        /*
         * PUT DLL FILE TO => C:\Program Files\Java\jdk1.8.0_171\jre\bin
         **/
        File file = new File(EXCEL_PATH);
        Excel.callExcelMacro(file, MACRO_NAME);
    }

    private static void test2(String cmd) {
//      cmd = "wscript C:/Users/sahir/Documents/Macros/myvbs.vbs";
        try {
            System.out.println("RUNNING COMMAND ===> " + cmd);
            Runtime.getRuntime().exec(cmd);
            System.out.println("=================================== ALL DONE ===========================================");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
