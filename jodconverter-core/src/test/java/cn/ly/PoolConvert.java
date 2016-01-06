package cn.ly;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.CancellationException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadPoolExecutor;

import org.apache.commons.io.FilenameUtils;
import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.StandardConversionTask;
import org.artofsolving.jodconverter.document.DefaultDocumentFormatRegistry;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.document.DocumentFormatRegistry;
import org.artofsolving.jodconverter.document.SimpleDocumentFormatRegistry;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.MockOfficeTask;
import org.artofsolving.jodconverter.office.OfficeException;
import org.artofsolving.jodconverter.office.OfficeManager;

import com.sun.star.document.UpdateDocMode;

import static org.testng.Assert.assertTrue;
import static org.testng.Assert.fail;

public class PoolConvert {
	// 设置任务执行超时时间， 分钟为单位
	private static final long TASK_EXECUTION_TIMEOUT = 1000 * 60 * 1L * 1;
	// 设置任务队列超时时间，分钟为单位 	// private static final long TASK_QUEUE_TIMEOUT = 1000 * 60 * 1L *
	private static final long TASK_QUEUE_TIMEOUT = 1000 * 60 * 1L * 1;
	public static void main(String[] args) throws IOException {

		final DefaultDocumentFormatRegistry dcfr = new DefaultDocumentFormatRegistry();

		int[] ports = {8810,8811,8812,8813,8814,8815,8816,8817};

		DefaultOfficeManagerConfiguration domc = new DefaultOfficeManagerConfiguration();
//		domc.setOfficeHome("/F:/Program Files/OpenOffice 4/").setPortNumbers(8810)
		domc.setOfficeHome("/F:/Program Files/OpenOffice 4/").setPortNumbers(ports)
				.setTemplateProfileDir(new File("F:/Program Files/OpenOffice 4/temp/"))
				.setTaskExecutionTimeout(TASK_EXECUTION_TIMEOUT)
				.setTaskQueueTimeout(TASK_QUEUE_TIMEOUT)
				.setMaxTasksPerProcess(ports.length-1);
		final OfficeManager officeManager = domc.buildOfficeManager();
		officeManager.start();

		// OfficeDocumentConverter converter = new
		// OfficeDocumentConverter(officeManager);
		Long star = System.currentTimeMillis();

		final Map<String,Object> loadProperties = new HashMap<String,Object>();
		loadProperties.put("Hidden", true);
		loadProperties.put("ReadOnly", true);
		loadProperties.put("UpdateDocMode", UpdateDocMode.QUIET_UPDATE);

		try {
			ExecutorService es = Executors.newFixedThreadPool(ports.length - 1);


			File dir = new File("d:/des/test/s/");
			File[] files = dir.listFiles();
			for (final File inputFile : files) {
				final DocumentFormat pdf = dcfr.getFormatByExtension("pdf");



				Runnable run = new Thread() {
					public void run() {
						File tFile = new File("d:/des/test/t/" + inputFile.getName() + ".pdf");
						OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
						try {
							converter.convert(inputFile, tFile);
						} catch (Exception e) {
							System.out.println("exception:"+e.getMessage());
						}

						System.out.printf("done.\n");
					}
				};

				es.execute(run);

				// if(tFile.exists()){
				// continue;
				// }else{
				// System.out.println(inputFile);
				//
				// converter.convert(inputFile,tFile, outFormat);
				// System.out.printf("done.\n");
				//
				// }
			}
		} finally {
			officeManager.stop();
		}
		System.out.println(System.currentTimeMillis() - star);
	}
}
