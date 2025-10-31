FROM node:18

# Install PDFtk, Ghostscript (optional), and qpdf (for appearance cleanup)
RUN apt-get update && apt-get install -y pdftk ghostscript qpdf openjdk-17-jdk curl && rm -rf /var/lib/apt/lists/*

# Fetch PDFBox (all-in-one jar)
RUN mkdir -p /opt && \
    curl -L -o /opt/pdfbox-app.jar https://repo1.maven.org/maven2/org/apache/pdfbox/pdfbox-app/2.0.29/pdfbox-app-2.0.29.jar

# Fetch iText (single all-in-one) for appearance regeneration (editable forms)
RUN mkdir -p /opt/itext && \
    curl -L -o /opt/itext/itext7-core.jar https://repo1.maven.org/maven2/com/itextpdf/itext7-core/7.2.5/itext7-core-7.2.5.jar

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
    javac -cp /opt/itext/itext7-core.jar -d /opt /app/RefreshAppearancesIText.java

# Expose port
EXPOSE 8080

# Start the application
CMD ["node", "index.js"]
