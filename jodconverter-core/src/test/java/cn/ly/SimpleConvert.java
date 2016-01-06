package cn.ly;

import com.sun.star.document.UpdateDocMode;
import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.StandardConversionTask;
import org.artofsolving.jodconverter.document.DefaultDocumentFormatRegistry;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeException;
import org.artofsolving.jodconverter.office.OfficeManager;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class SimpleConvert {
    // 设置任务执行超时时间
    private static final long TASK_EXECUTION_TIMEOUT = 120000L;  // 2 minutes
    // 设置任务队列超时时间
    private static final long TASK_QUEUE_TIMEOUT = 30000L;  // 30 seconds

    public static void main(String[] args) throws IOException {
        int[] ports = {8810};
        DefaultDocumentFormatRegistry dcfr = new DefaultDocumentFormatRegistry();
        DocumentFormat pdf = dcfr.getFormatByExtension("pdf");
        DefaultOfficeManagerConfiguration domc = new DefaultOfficeManagerConfiguration();
        domc.setOfficeHome("/F:/Program Files/OpenOffice 4/").setPortNumbers(ports)
                .setTemplateProfileDir(new File("F:/Program Files/OpenOffice 4/temp/"))
                .setTaskExecutionTimeout(TASK_EXECUTION_TIMEOUT)
                .setTaskQueueTimeout(TASK_QUEUE_TIMEOUT);
//				.setMaxTasksPerProcess(ports.length-1);
        OfficeManager officeManager = domc.buildOfficeManager();
        officeManager.start();

        Long star = System.currentTimeMillis();

        try {
            File dir = new File("d:/des/test/s/");
            File[] files = dir.listFiles();
            for (File sFile : files) {
                try {
                    File tFile = new File("d:/des/test/t/" + sFile.getName() + ".pdf");
                    OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
                    converter.convert(sFile, tFile);
                } catch (OfficeException e) {
                    e.printStackTrace();
                }
                System.out.println(sFile.getName()+" : done.");
            }
        } catch (Exception e) {
            System.out.println("exception:" + e.getMessage());
        } finally {
            officeManager.stop();
        }
        System.out.println(System.currentTimeMillis() - star);
    }
}
