import java.io.File;
import com.syncfusion.docio.*;

public class FindAndReplaceText {

	public static void main(String[] args) throws Exception {
		// Opens the input Word document.
		WordDocument document = new WordDocument(getDataDir("FindAndReplaceText_Template.docx"));
		// Finds all occurrences of a misspelled word and replaces with properly spelled word.
		document.replace("Cyles", "Cycles", true, true);
		// Saves the resultant file in the given path.
		document.save("Result.docx");
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
