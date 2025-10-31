import java.io.File;
import java.util.List;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.interactive.form.*;
import org.apache.pdfbox.cos.COSDictionary;
import org.apache.pdfbox.cos.COSName;

public class RefreshAppearances {
  public static void main(String[] args) throws Exception {
    if (args.length < 2) {
      System.err.println("Usage: java -cp /opt/pdfbox-app.jar:/opt RefreshAppearances in.pdf out.pdf");
      System.exit(1);
    }
    File in = new File(args[0]);
    File out = new File(args[1]);
    try (PDDocument doc = PDDocument.load(in)) {
      // Bypass PDFBox default fixup (which triggers full refreshAppearances)
      COSDictionary acroDict = (COSDictionary) doc.getDocumentCatalog().getCOSObject().getDictionaryObject(COSName.ACRO_FORM);
      PDAcroForm form = acroDict != null ? new PDAcroForm(doc, acroDict) : null;
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
        PDRadioButton rb = (PDRadioButton) field;
        String v = rb.getValue();
        if (v != null) rb.setValue(v); // re-apply same value to rebuild /AP
      } else if (field instanceof PDCheckBox) {
        PDCheckBox cb = (PDCheckBox) field;
        String v = cb.getValue();
        if (v == null || v.equalsIgnoreCase("Off")) {
          cb.unCheck();
        } else {
          cb.setValue(v); // e.g., "Yes" or custom on-value
        }
      } // skip text fields etc.
    } catch (Exception ignore) {
      // keep going even if a particular field fails
    }
  }
}


