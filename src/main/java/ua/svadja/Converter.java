package ua.svadja;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Random;

import com.sun.star.beans.PropertyValue;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lib.uno.adapter.ByteArrayToXInputStreamAdapter;
import com.sun.star.lib.uno.adapter.OutputStreamToXOutputStreamAdapter;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import ooo.connector.BootstrapConnector;
import ooo.connector.server.OOoServer;

public class Converter {
	private String loFolder;

	/**
	 * @param loFolder
	 *            - Libre Office program folder. Example: "c:/Program
	 *            Files/LibreOffice 5/program/"
	 */
	public Converter(String loFolder) {
		super();
		this.loFolder = loFolder;
	}

	public synchronized OutputStream convertDocxToPdf(InputStream in) {
		BootstrapConnector bootstrapConnector = null;
		OutputStream out = new ByteArrayOutputStream();
		try {
			String sPipeName = "uno" + Long.toString((new Random()).nextLong() & 0x7fffffffffffffffL);
			OOoServer server = createLOServerWithPipe(this.loFolder, sPipeName);
			bootstrapConnector = new BootstrapConnector(server);
			XComponentContext xContext = bootstrapConnector.connect("",
					"uno:pipe,name=" + sPipeName + ";urp;StarOffice.ComponentContext");
			XMultiComponentFactory xMCF = xContext.getServiceManager();
			Object oDesktop = xMCF.createInstanceWithContext("com.sun.star.frame.Desktop", xContext);
			XDesktop xDesktop = UnoRuntime.queryInterface(XDesktop.class, oDesktop);
			XComponentLoader xCompLoader = UnoRuntime.queryInterface(XComponentLoader.class, xDesktop);

			PropertyValue[] propertyValue = new PropertyValue[2];
			propertyValue[0] = new PropertyValue();
			propertyValue[0].Name = "InputStream";
			propertyValue[0].Value = new ByteArrayToXInputStreamAdapter(InputStreamToByteArray(in));
			propertyValue[1] = new PropertyValue();
			propertyValue[1].Name = "Hidden";
			propertyValue[1].Value = new Boolean(true);
			XComponent documentComponent = xCompLoader.loadComponentFromURL("private:stream", "_blank", 0,
					propertyValue);

			propertyValue = new PropertyValue[2];
			propertyValue[0] = new PropertyValue();
			propertyValue[0].Name = "OutputStream";
			propertyValue[0].Value = new OutputStreamToXOutputStreamAdapter(out);
			propertyValue[1] = new PropertyValue();
			propertyValue[1].Name = "FilterName";
			propertyValue[1].Value = "writer_pdf_Export";
			XStorable xstorable = UnoRuntime.queryInterface(XStorable.class, documentComponent);
			xstorable.storeToURL("private:stream", propertyValue);

		} catch (BootstrapException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (bootstrapConnector != null) {
				bootstrapConnector.disconnect();
			}
		}

		return out;
	}

	private OOoServer createLOServerWithPipe(String loFolder, String pipeName) {
		List<String> options = OOoServer.getDefaultOOoOptions();
		options.add("--accept=pipe,name=" + pipeName + ";urp;");
		options.add("--nofirststartwizard");
		options.add("--headless");
		return new OOoServer(this.loFolder, options);
	}

	private byte[] InputStreamToByteArray(InputStream is) throws IOException {
		ByteArrayOutputStream buffer = new ByteArrayOutputStream();
		int nRead;
		byte[] data = new byte[16384];
		while ((nRead = is.read(data, 0, data.length)) != -1) {
			buffer.write(data, 0, nRead);
		}
		buffer.flush();
		return buffer.toByteArray();

	}

}
