import java.io.*;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.drawing.ColorSupport;

public class FindAndHighlightText {

	public static void main(String[] args) throws Exception {
		// Opens the input Word document.
		WordDocument document = new WordDocument(getDataDir("FindAndHighlightText_Template.docx"), FormatType.Docx);
		// Finds all occurrence of the text in the Word document.
		TextSelection[] textSelections = document.findAll("Adventure", true, true);
		for (int i = 0; i < textSelections.length; i++) {
			// Sets the highlight color for the searched text as Yellow.
			textSelections[(int) i].getAsOneRange().getCharacterFormat().setHighlightColor(ColorSupport.getYellow());
		}
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
