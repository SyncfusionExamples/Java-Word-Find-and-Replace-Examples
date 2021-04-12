import java.io.*;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;
import com.syncfusion.javahelper.system.io.*;
import com.syncfusion.javahelper.system.xml.*;

public class ReplacTextWithTable {

	public static void main(String[] args) throws Exception {
		// Opens the input Word document.
		WordDocument document = new WordDocument(getDataDir("ReplaceTextWithTable_Template.docx"));
		// Creates a new table.
		WTable table = new WTable(document);
		table.resetCells(1, 6);
		table.get(0, 0).setWidth(52f);
		table.get(0, 0).addParagraph().appendText("Supplier ID");
		table.get(0, 1).setWidth(128f);
		table.get(0, 1).addParagraph().appendText("Company Name");
		table.get(0, 2).setWidth(70f);
		table.get(0, 2).addParagraph().appendText("Contact Name");
		table.get(0, 3).setWidth(92f);
		table.get(0, 3).addParagraph().appendText("Address");
		table.get(0, 4).setWidth(66.5f);
		table.get(0, 4).addParagraph().appendText("City");
		table.get(0, 5).setWidth(56f);
		table.get(0, 5).addParagraph().appendText("Country");
		// Imports data to the table.
		importDataToTable(table);
		// Applies the built-in table style (Medium Shading 1 Accent 1) to the table.
		table.applyStyle(BuiltinTableStyle.MediumShading1Accent1);
		TextBodyPart bodyPart = new TextBodyPart(document);
		bodyPart.getBodyItems().add(table);
		// Replaces the table placeholder text with a new table.
		document.replace("[Suppliers table]", bodyPart, true, true, true);
		// Saves the resultant file in the given path.
		document.save("Result.docx");
	}

	private static void importDataToTable(WTable table) throws Exception {
		FileStreamSupport fs = new FileStreamSupport(getDataDir("Suppliers.xml"), FileMode.Open, FileAccess.Read);
		XmlReaderSupport reader = XmlReaderSupport.create(fs);
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (reader.getLocalName() != "SuppliersList")
			throw new Exception(StringSupport.concat("Unexpected xml tag ", reader.getLocalName()));
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		while (reader.getLocalName() != "SuppliersList") {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch (reader.getLocalName()) {
				case "Suppliers":
					WTableRow tableRow = table.addRow(true);
					importDataToRow(reader, tableRow);
					break;
				}
			} else {
				reader.read();
				if ((reader.getLocalName() == "SuppliersList") && reader.getNodeType() == XmlNodeType.EndElement)
					break;
			}
		}
		reader.close();
		fs.close();
	}

	private static void importDataToRow(XmlReaderSupport reader, WTableRow tableRow) throws Exception {
		if (reader == null)
			throw new Exception("reader");
		while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
			reader.read();
		if (reader.getLocalName() != "Suppliers")
			throw new Exception(StringSupport.concat("Unexpected xml tag ", reader.getLocalName()));
		reader.read();
		while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
			reader.read();
		while (reader.getLocalName() != "Suppliers") {
			if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) {
				switch (reader.getLocalName()) {
				case "SupplierID":
					tableRow.getCells().get(0).addParagraph().appendText(reader.readContentAsString());
					break;
				case "CompanyName":
					tableRow.getCells().get(1).addParagraph().appendText(reader.readContentAsString());
					break;
				case "ContactName":
					tableRow.getCells().get(2).addParagraph().appendText(reader.readContentAsString());
					break;
				case "Address":
					tableRow.getCells().get(3).addParagraph().appendText(reader.readContentAsString());
					break;
				case "City":
					tableRow.getCells().get(4).addParagraph().appendText(reader.readContentAsString());
					break;
				case "Country":
					tableRow.getCells().get(5).addParagraph().appendText(reader.readContentAsString());
					break;
				default:
					reader.skip();
					break;
				}
			} else {
				reader.read();
				if ((reader.getLocalName() == "Suppliers") && reader.getNodeType() == XmlNodeType.EndElement)
					break;
			}
		}
	}

	/**
	 * Get the file path
	 * 
	 * @param path specifies the file path
	 */
	private static String getDataDir(String path) {
		File dir = new File(System.getProperty("user.dir"));
		if (!(dir.toString().endsWith("Java-Word-Find-and-Replace-Examples")))
			dir = dir.getParentFile();
		dir = new File(dir, "resources");
		dir = new File(dir, path);
		if (dir.isDirectory() == false)
			dir.mkdir();
		return dir.toString();
	}
}
