import java.io.*;
import java.util.regex.Pattern;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;
import com.syncfusion.javahelper.system.text.regularExpressions.MatchSupport;

public class ReplaceTextWithMergeField {

	public static void main(String[] args) throws Exception {
		// Opens the input Word document.
		WordDocument document = new WordDocument(getDataDir("ReplaceWithMergeFileds_Template.docx"));
		// Finds all the placeholder text enclosed within '«' and '»' in the Word document.
		TextSelection[] textSelections = document.findAll(
				Pattern.compile(MatchSupport.trimPattern("«([(?i)image(?-i)]*:*[a-zA-Z0-9 ]*:*[a-zA-Z0-9 ]+)»")));
		String[] searchedPlaceholders = new String[textSelections.length];
		for (int i = 0; i < textSelections.length; i++) {
			searchedPlaceholders[(int) i] = textSelections[(int) i].getSelectedText();
		}
		for (int i = 0; i < searchedPlaceholders.length; i++) {
			WParagraph paragraph = new WParagraph(document);
			// Replaces the placeholder text enclosed within '«' and '»' with desired merge field.
			paragraph.appendField(
					StringSupport.trimEnd(StringSupport.trimStart(searchedPlaceholders[i], '«'), '»'),
					FieldType.FieldMergeField);
			TextSelection newSelection = new TextSelection(paragraph, 0, paragraph.getItems().getCount());
			TextBodyPart bodyPart = new TextBodyPart(document);
			bodyPart.getBodyItems().add(paragraph);
			document.replace(searchedPlaceholders[(int) i], bodyPart, true, true, true);
		}
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
