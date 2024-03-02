# Use the Python 3.9 image as the base
FROM python:3.9

# Set the working directory inside the container
WORKDIR /app

# Copy the current directory (containing all files) into the container's /app directory
COPY . /app/

# Install Python dependencies listed in requirements.txt without caching
RUN pip install --no-cache-dir -r requirements.txt

# Expose port 80 to allow communication with other services
EXPOSE 80

# Set an environment variable named NAME with the value 'gaghiel_env'
ENV NAME gaghiel_env

# Specify the command to run when the container starts
CMD ["python", "main.py"]
