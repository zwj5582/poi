package com.excel;

import com.excel.poi.reader.DefaultPoiPropertiesReader;
import com.excel.poi.PoiExporter;
import com.excel.poi.exception.TooManyDataExportException;
import lombok.*;
import ognl.OgnlException;
import org.apache.poi.hssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class 	PoiApplication {

	@Getter@Setter
	@AllArgsConstructor@NoArgsConstructor
	public static class Patient{
		private String name;
		private Integer age;
		private String sex;
		private Date date;
	}

	@Getter@Setter
	@AllArgsConstructor@NoArgsConstructor
	public static class Task{
		private String address;
		private Patient patient;
	}

	@SuppressWarnings("unchecked")
	public static void main(String[] args) throws IOException, OgnlException, TooManyDataExportException, InstantiationException, IllegalAccessException, ClassNotFoundException {
		String filePath="D:\\Dropbox\\";//文件路径

		File file = new File(filePath);
		if (!file.exists()) {
			boolean mkdirs = file.mkdirs();
			System.out.println(mkdirs);
		}
		filePath+="sample.xls";

		List<Task> entityList = new ArrayList<Task>();

		entityList.add(new Task("hfjkdshfjks",new Patient("钟文杰", 99, "男",new Date())));
		entityList.add(new Task("GDFGDF",new Patient("钟文杰11", 88, "男1",new Date())));
		entityList.add(new Task("hfjkdshfjkD",new Patient("钟文杰22", 111, "男2",new Date())));
		entityList.add(new Task("KSKKDJSDK",new Patient("钟文杰33", 222, "男3",new Date())));
		entityList.add(new Task("hfjkdshfjksk",new Patient("钟文杰44", 333, "男4",new Date())));
		entityList.add(new Task("DFDFDSF",new Patient("钟文杰55", 444, "男5",new Date())));

		PoiExporter exporter = new PoiExporter(new DefaultPoiPropertiesReader());

		HSSFWorkbook workbook = exporter.createExporter(new WorkBookEntity(), entityList);



		FileOutputStream out = new FileOutputStream(filePath);

		workbook.write(out);//保存Excel文件


		out.close();//关闭文件流
		System.out.println("OK!");

	}
}
