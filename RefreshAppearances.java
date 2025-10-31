import java.io.File;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.interactive.form.PDAcroForm;

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
        form.refreshAppearances();
      }
      doc.save(out);
    }
  }
}


