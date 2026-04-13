# Use a lightweight Python 3.11 image
FROM python:3.11-slim

# Hugging Face requires running as a non-root user
RUN useradd -m -u 1000 user
USER user

# Set environment variables for the user
ENV HOME=/home/user \
    PATH=/home/user/.local/bin:$PATH

# Set the working directory
WORKDIR $HOME/app

# Copy all your project folders into the container
COPY --chown=user . .

# Install requirements from BOTH folders
RUN pip install --no-cache-dir -r backend/requirements.txt
RUN pip install --no-cache-dir -r frontend/requirements.txt

# Expose the specific port Hugging Face looks for
EXPOSE 7860

# Tell Docker to run app.py from the frontend folder
CMD ["streamlit", "run", "frontend/app.py", "--server.port=7860", "--server.address=0.0.0.0"]
