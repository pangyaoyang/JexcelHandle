package com.pangyy.jexcelHandle.common;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteFile {
	
	// 向文件写入内容(输出流)
	public static void writeFile(String fileName, String context) {
		byte bt[] = new byte[context.length() + 100];
		bt = context.getBytes();

		try {
			FileOutputStream in = new FileOutputStream(fileName);
			try {
				in.write(bt, 0, bt.length);
				in.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}

	// 向文件写入内容(输出流)
	public static void writeFile(String fileName, StringBuffer context) {
		byte bt[] = new byte[context.toString().length() + 100];
		bt = context.toString().getBytes();

		try {
			FileOutputStream in = new FileOutputStream(fileName);
			try {
				in.write(bt, 0, bt.length);
				in.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
}
