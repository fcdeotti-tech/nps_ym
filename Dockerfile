# Usa uma imagem oficial e leve do Python
FROM python:3.10-slim

# Define a pasta de trabalho dentro do servidor
WORKDIR /app

# Copia o arquivo de bibliotecas primeiro e instala
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia o restante do seu projeto (script e source) para o servidor
COPY . .

# Comando exato para ligar o Streamlit usando a porta dinâmica do Google ($PORT)
CMD ["streamlit", "run", "script/app.py", "--server.port=8080", "--server.address=0.0.0.0", "--server.enableCORS=false", "--server.enableXsrfProtection=false"]
