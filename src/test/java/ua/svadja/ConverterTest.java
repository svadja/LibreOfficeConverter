package ua.svadja;

import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import java.io.BufferedInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

import org.junit.Before;
import org.junit.Test;

public class ConverterTest {
	private Converter converter;

	@Before
	public void setUp() throws Exception {
		converter = new Converter("c:/Program Files/LibreOffice 5/program/");
	}

	@Test
	public void testConvert() {
		InputStream in;
		try {
			in = new BufferedInputStream(new FileInputStream("src/test/resources/test.docx"));
			ByteArrayOutputStream out = (ByteArrayOutputStream) converter.convertDocxToPdf(in);
			assertNotNull(out);
			assertTrue(out.size()>0);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}