import java.io.File;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        String cmd = "wscript " + new File(Main.class.getClassLoader().getResource("myvbs.vbs").getFile()).getAbsolutePath();
        test2(cmd);
    }

    private static void test1() {
        /*
         * PUT DLL FILE TO => C:\Program Files\Java\jdk1.8.0_171\jre\bin
         **/
        File file = new File("C:\\Users\\sahir\\Documents\\Macros\\MacroWS.xlsm");
        String macroName = "NewMacro2";
        Excel.callExcelMacro(file, macroName);
    }

    private static void test2(String cmd) {
//      cmd = "wscript C:/Users/sahir/Documents/Macros/myvbs.vbs";
        try {
            Runtime.getRuntime().exec(cmd);
            System.out.println("=================================== ALL DONE ===========================================");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
