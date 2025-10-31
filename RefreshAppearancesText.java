import java.io.File;
import java.util.Map;
import com.itextpdf.forms.PdfAcroForm;
import com.itextpdf.forms.fields.PdfFormField;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.PdfWriter;

public class RefreshAppearancesIText {
  public static void main(String[] args) throws Exception {
    if (args.length < 2) {
      System.err.println("Usage: java -cp itext-jars RefreshAppearancesIText in.pdf out.pdf");
      System.exit(1);
    }
    File in = new File(args[0]);
    File out = new File(args[1]);
    try (PdfDocument pdf = new PdfDocument(new PdfReader(in), new PdfWriter(out))) {
      PdfAcroForm form = PdfAcroForm.getAcroForm(pdf, true);
      form.setGenerateAppearance(true);
      // Re-apply existing value to trigger appearance generation
      for (Map.Entry<String, PdfFormField> e : form.getFormFields().entrySet()) {
        PdfFormField f = e.getValue();
        String v = f.getValueAsString();
        if (v != null) {
          try { f.setValue(v, true); } catch (Exception ignore) {}
        }
      }
      // Editable: do not flatten
    }
  }
}


