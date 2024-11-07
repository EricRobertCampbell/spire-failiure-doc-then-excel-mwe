# Start with the ubuntu:22.04 base image
FROM ubuntu:22.04

# Installations
RUN apt-get update && \
    apt-get install -y g++ && \
    ln -sf g++ /usr/bin/c++
RUN apt-get install -y libicu-dev
RUN apt-get update && \
    apt-get install -y python3-dev
RUN apt-get update && \
    apt-get install -y python3-pip
RUN apt-get update && \
    apt-get install -y unixodbc-dev

# Set working directory
WORKDIR /app

# Copy files
COPY ./fonts/* /usr/share/fonts/
COPY MainLogo.png .
COPY Spire.Doc-12.7.1-py3-none-manylinux1_x86_64.whl .
COPY Spire.Xls-14.7.3-py3-none-manylinux1_x86_64.whl .

# Install Spire products
RUN pip install Spire.Doc-12.7.1-py3-none-manylinux1_x86_64.whl 
RUN pip install Spire.Xls-14.7.3-py3-none-manylinux1_x86_64.whl

# Copy the Python scripts 
COPY crash_change.py ./

# Run crash_change.py
CMD ["python3", "crash_change.py"]
