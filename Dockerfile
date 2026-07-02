# --- Estágio de Compilação (Builder) ---
FROM python:3.11-slim AS builder

WORKDIR /app

# Cria o virtual environment para isolar dependências compiladas
RUN python -m venv /opt/venv
ENV PATH="/opt/venv/bin:$PATH"

# Copia apenas o requirements.txt primeiro para aproveitar o cache de camadas do Docker
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt


# --- Estágio de Execução Final (Runner) ---
FROM python:3.11-slim AS runner

WORKDIR /app

# Instala o Chromium e o driver WebDriver correspondente, além de limpar o cache do apt
RUN apt-get update && apt-get install -y --no-install-recommends \
    chromium \
    chromium-driver \
    && rm -rf /var/lib/apt/lists/*

# Cria um link simbólico para que o Selenium encontre o Chromium como "google-chrome"
RUN ln -s /usr/bin/chromium /usr/bin/google-chrome

# Copia o virtual environment completo do estágio de build
COPY --from=builder /opt/venv /opt/venv
ENV PATH="/opt/venv/bin:$PATH"

# Cria um grupo e um usuário com ID 10001
RUN groupadd -g 10001 appuser && \
    useradd -r -u 10001 -g appuser -d /app appuser

# Cria a estrutura de diretórios obrigatórios do projeto
RUN mkdir -p /app/dados/temp_downloads /app/dados/xml_materiais_quirius /app/dados/historico

# Copia os códigos-fonte da aplicação
COPY src/ /app/src/

# Altera as permissões de gravação da pasta app para o usuário criado
RUN chown -R appuser:appuser /app

# Executa o container com usuário sem privilégios administrativos (Segurança)
USER appuser

# Configurações de ambiente de tempo de execução
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    SE_OFFLINE=true \
    HOME=/app
