package cn.ly;

import com.sun.star.document.UpdateDocMode;
import org.apache.commons.io.FilenameUtils;
import org.artofsolving.jodconverter.StandardConversionTask;
import org.artofsolving.jodconverter.document.DefaultDocumentFormatRegistry;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class PoolConvert3 {
	static final DefaultDocumentFormatRegistry dcfr = new DefaultDocumentFormatRegistry();
	static final DocumentFormat pdf = dcfr.getFormatByExtension("pdf");
	static final Map<String,Object> loadProperties = new HashMap<String,Object>();
	static {
		loadProperties.put("Hidden", true);
		loadProperties.put("ReadOnly", true);
		loadProperties.put("UpdateDocMode", UpdateDocMode.QUIET_UPDATE);
	}
	static int[] ports = {8810,8811,8812,8813,8814};

	static DefaultOfficeManagerConfiguration domc = new DefaultOfficeManagerConfiguration();
	static {
		domc.setOfficeHome("/F:/Program Files/OpenOffice 4/").setPortNumbers(ports)
		.setTemplateProfileDir(new File("F:/Program Files/OpenOffice 4/temp/"));
	}
	static final OfficeManager officeManager = domc.buildOfficeManager();
	static {
		officeManager.start();
	}
	public static void main(String[] args) throws IOException {





		// OfficeDocumentConverter converter = new
		// OfficeDocumentConverter(officeManager);
		Long star = System.currentTimeMillis();


		ExecutorService es = null;

		try {
			es = Executors.newFixedThreadPool(3);


			File dir = new File("d:/des/test/s/");
			File[] files = dir.listFiles();
			for (final File inputFile : files) {
				Runnable run = new convertTestPoolInner(inputFile);
				es.execute(run);
			}
		} catch (Exception e){
			e.printStackTrace();

		}finally {
			officeManager.stop();
			es.shutdown();
		}
		System.out.println(System.currentTimeMillis() - star);
	}
}

class testInner implements Runnable{
	File sFile = null;

	public testInner(File sFile) {
		this.sFile = sFile;
	}

	public void run() {
		System.out.println(sFile.getName());
		File tFile = new File("d:/des/test/t/" + sFile.getName() + ".pdf");
		StandardConversionTask conversionTask = new StandardConversionTask(sFile, tFile, PoolConvert1.pdf);
		conversionTask.setDefaultLoadProperties(PoolConvert1.loadProperties);
		conversionTask.setInputFormat(PoolConvert1.dcfr.getFormatByExtension(FilenameUtils.getExtension(sFile.getName())));
		PoolConvert1.officeManager.execute(conversionTask);

		System.out.printf("done.\n");

	}
}
