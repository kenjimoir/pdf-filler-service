import java.io.File;
import java.util.List;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.interactive.form.*;

public class RefreshAppearances {
  public static void main(String[] args) throws Exception {
    if (args.length < 2) {
      System.err.println("Usage: java -cp /opt/pdfbox-app.jar:/opt RefreshAppearances in.pdf out.pdf");
      System.exit(1);
    }
    File in = new File(args[0]);
    File out = new File(args[1]);
    try (PDDocument doc = PDDocument.load(in)) {
      PDAcroForm form = doc.getDocumentCatalog().getAcroForm();
      if (form != null) {
        // Only rebuild appearances for buttons (checkbox/radio) to avoid font encoding issues
        refreshButtonAppearances(form);
      }
      doc.save(out);
    }
  }

  private static void refreshButtonAppearances(PDAcroForm form) {
    for (PDField field : form.getFields()) {
      refreshField(field);
    }
  }

  private static void refreshField(PDField field) {
    try {
      if (field instanceof PDNonTerminalField) {
        List<PDField> kids = ((PDNonTerminalField) field).getChildren();
        if (kids != null) for (PDField k : kids) refreshField(k);
      } else if (field instanceof PDRadioButton) {
        field.constructAppearances();
      } else if (field instanceof PDCheckBox) {
        field.constructAppearances();
      } // skip text fields etc.
    } catch (Exception ignore) {
      // keep going even if a particular field fails
    }
  }
}


