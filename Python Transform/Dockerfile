# Use the official Python 3.7 image from Docker Hub
FROM python:3.8

# Set a working directory inside the container
WORKDIR /app

# Copy your Python script into the container
COPY report_form_1.py .
COPY report_form_2.py .
COPY report_form_3.py .
COPY report_form_4.py .
COPY report_form_1_debug.py .

# Copy the specific CSV files from your host machine into the container
COPY Reports/1.csv .
COPY Reports/2.csv .
COPY Reports/3.csv .
COPY Reports/4.csv .

# Install any dependencies your script requires
# You might need to add specific dependencies if your script uses non-standard libraries
RUN pip install pandas
RUN pip install xlsxwriter
