import java.io.*;
import java.util.regex.Pattern;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;
import com.syncfusion.javahelper.system.text.regularExpressions.MatchSupport;

public class ReplaceTextWithDocument {

	public static void main(String[] args) throws Exception {
		// Opens the Word template document
		WordDocument document = new WordDocument(getDataDir("ReplaceTextWithDocument_Template.docx"), FormatType.Docx);
		TextSelection[] textSelections = document.findAll(Pattern.compile(MatchSupport.trimPattern("\\[(.*)\\]")));
		for (int i = 0; i < textSelections.length; i++) {
			WordDocument subDocument = new WordDocument(getDataDir(
					StringSupport.trimEnd(StringSupport.trimStart(textSelections[i].getSelectedText(), '['), ']')
							+ ".docx"),
					FormatType.Docx);
			document.replace(textSelections[(int) i].getSelectedText(), subDocument, true, true);
			subDocument.close();
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
