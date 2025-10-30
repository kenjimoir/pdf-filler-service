FROM node:18

# Install PDFtk
RUN apt-get update && apt-get install -y pdftk

# Set working directory
WORKDIR /app

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm install

# Copy source code
COPY . .

# Expose port
EXPOSE 8080

# Start the application
CMD ["node", "index.js"]
