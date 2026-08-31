FROM python:3.11-slim

# LibreOffice is what finalize.py shells out to for the "opens clean in Excel" pass.
# Without it the kit still runs (per README) but skips that normalize step.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn

COPY . .
RUN mkdir -p inputs/current inputs/lastweek inputs/trades outputs

EXPOSE 8000

# 1 worker, several threads: JOB state in webapp.py is in-memory and per-process,
# so multiple gunicorn workers would each show a different "is it running" status.
CMD ["gunicorn", "-w", "1", "--threads", "4", "--timeout", "300", "-b", "0.0.0.0:8000", "webapp:app"]
