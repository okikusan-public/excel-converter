FROM --platform=linux/amd64 node:18-slim

# Install Java JDK and build tools
RUN apt-get update && apt-get install -y \
    default-jdk \
    build-essential \
    python3 \
    fonts-ipafont-gothic \
    fonts-ipafont-mincho \
    fonts-ipaexfont \
    fonts-noto-cjk \
    locales \
    && locale-gen ja_JP.UTF-8 \
    && dpkg-reconfigure locales \
    # Create symlink from python3 to python for compatibility with node-gyp and other build tools
    && ln -sf /usr/bin/python3 /usr/bin/python \
    && rm -rf /var/lib/apt/lists/*

# Set environment variables
ENV LANG=ja_JP.UTF-8
ENV LANGUAGE=ja_JP.UTF-8
ENV LC_ALL=ja_JP.UTF-8
ENV JAVA_HOME=/usr/lib/jvm/default-java

# Create app directory
WORKDIR /app

# Copy package.json and package-lock.json
COPY package*.json ./

# Install dependencies
RUN npm install

# Copy source code
COPY . .

# Build TypeScript code
RUN npm run build

# Create input and output directories
RUN mkdir -p /app/input /app/output /app/aspose.cells /app/temp

# Set volumes
VOLUME ["/app/input", "/app/output", "/app/aspose.cells", "/app/temp"]

# Run the application
CMD ["npm", "start"]
