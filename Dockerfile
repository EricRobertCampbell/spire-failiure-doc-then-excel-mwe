# Start with the ubuntu:22.04 base image
FROM ubuntu:22.04

# Install necessary packages and Python 3.9.16
ARG DEBIAN_FRONTEND=noninteractive
RUN apt-get update && \
    apt-get install -y software-properties-common && \
    add-apt-repository -y ppa:deadsnakes/ppa && \
    apt-get update && \
    apt-get install -y python3.9 python3.9-venv && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

# c++ environment from the dockerfile sent over by spire
RUN apt-get update && \
    apt-get install -y g++ && \
    ln -sf g++ /usr/bin/c++
RUN apt-get install -y libicu-dev
# This had to be changed because we are on python 3.9.16
# RUN apt-get update && \
#     apt-get install -y python3-dev
# RUN apt-get update && \
#     apt-get install -y python3-pip
RUN apt-get update && \
    apt-get install -y unixodbc-dev

# Set the default Python 3.9 as 'python3'
RUN update-alternatives --install /usr/bin/python3 python3 /usr/bin/python3.9 1

# Set working directory
WORKDIR /app

COPY requirements.txt .
COPY MainLogo.png .
RUN mkdir results

# Create a virtual environment and install requirements
RUN python3 -m venv venv && \
    . venv/bin/activate && \
    pip install --no-cache-dir -r requirements.txt

# Copy the Python scripts into the container
COPY crash.py crash_change.py ./

# Activate the virtual environment and run crash_change.py
CMD ["sh", "-c", ". venv/bin/activate && python3 crash_change.py"]
