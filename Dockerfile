FROM python:3.12-slim

RUN apt-get update \
    && apt-get install -y --no-install-recommends fonts-nanum \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

EXPOSE 10000
CMD ["sh", "-c", "streamlit run app.py --server.address 0.0.0.0 --server.port ${PORT:-10000} --server.headless true"]
