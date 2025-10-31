FROM node:18

# Install PDFtk, Ghostscript (optional), and qpdf (for appearance cleanup)
RUN apt-get update && apt-get install -y pdftk ghostscript qpdf openjdk-17-jdk curl && rm -rf /var/lib/apt/lists/*

# Fetch PDFBox (all-in-one jar)
RUN mkdir -p /opt && \
    curl -L -o /opt/pdfbox-app.jar https://repo1.maven.org/maven2/org/apache/pdfbox/pdfbox-app/2.0.29/pdfbox-app-2.0.29.jar

# Fetch iText jars for appearance regeneration (editable forms)
RUN mkdir -p /opt/itext && \
    curl -L -o /opt/itext/itext-kernel.jar https://repo1.maven.org/maven2/com/itextpdf/itextkernel/8.0.2/itextkernel-8.0.2.jar && \
    curl -L -o /opt/itext/itext-forms.jar  https://repo1.maven.org/maven2/com/itextpdf/itextforms/8.0.2/itextforms-8.0.2.jar && \
    curl -L -o /opt/itext/commons-io.jar   https://repo1.maven.org/maven2/commons-io/commons-io/2.11.0/commons-io-2.11.0.jar

# Set working directory
WORKDIR /app

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm install

# Copy source code
COPY . .

# Compile the small appearance refresh helpers (fail build if compile fails)
RUN javac -cp /opt/pdfbox-app.jar -d /opt /app/RefreshAppearances.java && \
    javac -cp /opt/itext/itext-kernel.jar:/opt/itext/itext-forms.jar:/opt/itext/commons-io.jar -d /opt /app/RefreshAppearancesIText.java

# Expose port
EXPOSE 8080

# Start the application
CMD ["node", "index.js"]
