import java.io.*;
import com.syncfusion.docio.*;

public class FindAndReplaceSeveralParagraphs {

	public static void main(String[] args) throws Exception {
		// Opens the input Word document.
		WordDocument document = new WordDocument(getDataDir("FindAndReplaceSeveralParagraphs_Template.docx"),
				FormatType.Docx);
		WordDocument subDocument = new WordDocument(getDataDir("Source.docx"), FormatType.Docx);
		// Gets the content from another Word document to replace.
		TextBodyPart replacePart = new TextBodyPart(subDocument);
		for (Object bodyItem_tempObj : subDocument.getLastSection().getBody().getChildEntities()) {
			TextBodyItem bodyItem = (TextBodyItem) bodyItem_tempObj;
			replacePart.getBodyItems().add(bodyItem.clone());
		}
		String placeholderText = "Suppliers/Vendors of Northwind" + "Customers of Northwind"
				+ "Employee details of Northwind traders" + "The product information" + "The inventory details"
				+ "The shippers" + "Purchase Order transactions" + "Sales Order transaction" + "Inventory transactions"
				+ "Invoices" + "[end replace]";
		// Finds the text that extends to several paragraphs and replaces it with
		// desired content.
		document.replaceSingleLine(placeholderText, replacePart, false, false);
		subDocument.close();
		// Saves the Word document
		document.save("Result.docx");
		// Closes the document
		document.close();
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
