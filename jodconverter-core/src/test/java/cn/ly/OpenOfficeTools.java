package cn.ly;

import java.io.File;

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;

/**
 * 文件转换工具类 主要能将offcie txt类型文件转化成 pdf  *   * @author xzy 2013-6-8下午03:54:50  *
 */
public class OpenOfficeTools {
    private static OfficeManager officeManager;
    private static OpenOfficeTools instance = new OpenOfficeTools();
// 设置任务执行超时时间， 分钟为单位
    private static final long TASK_EXECUTION_TIMEOUT = 1000 * 60 * 1L * 1;
// 设置任务队列超时时间，分钟为单位 	// private static final long TASK_QUEUE_TIMEOUT = 1000 * 60 * 1L *
	private static final long TASK_QUEUE_TIMEOUT = 1000 * 60 * 1L * 1;
    public static OpenOfficeTools getInstance() { 		return instance; 	}

    public static boolean convert2PDF(File inputFile,File pdfFile) {
        startService();
        OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
        try {
            converter.convert(inputFile, pdfFile);
        } catch (Exception e) {
            System.out.println("exception:"+e.getMessage());
        }
    	stopService();
        return pdfFile.isFile();
    }
    /** 	 * 启动openoffice服务 	 */
    private static void startService() {
        DefaultOfficeManagerConfiguration domc = new DefaultOfficeManagerConfiguration();
        //		setProcessManager
        try {
            domc.setOfficeHome("/F:/Program Files/OpenOffice 4/").setPortNumbers(8100, 8101, 8102, 8103)
                .setTemplateProfileDir(new File("F:/Program Files/OpenOffice 4/temp/"))
                .setTaskExecutionTimeout(TASK_EXECUTION_TIMEOUT).setTaskQueueTimeout(TASK_QUEUE_TIMEOUT);
            domc.setMaxTasksPerProcess(3);
            officeManager = domc.buildOfficeManager();

            officeManager.start(); // 启动服务
        } catch (Exception ce) { //
            ce.printStackTrace();
        }
    }
    /** 	 * 关闭openoffice服务 	 */
    private static void stopService() {
        if (officeManager != null) {
            officeManager.stop();
        }
    }
}