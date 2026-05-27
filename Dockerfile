# Container image for the Bizzdesign fork of GongRzhe's
# Office-PowerPoint-MCP-Server. Designed to run as an AWS Lambda container
# function fronted by the AWS Lambda Web Adapter, which translates API
# Gateway invocations into HTTP requests against the FastMCP HTTP transport
# listening on PORT.
#
# Differs from the upstream Smithery Dockerfile (Alpine + stdio entrypoint)
# in two ways:
#  - Debian slim base, since python-pptx + Pillow pull in native deps that
#    are flaky to build on Alpine.
#  - HTTP transport on port 8000 + lambda-adapter extension, so the same
#    image runs locally (`docker run -p 8000:8000 ...`) and on Lambda.
FROM public.ecr.aws/awsguru/aws-lambda-adapter:0.9.0 AS adapter

FROM python:3.12-slim

COPY --from=adapter /lambda-adapter /opt/extensions/lambda-adapter

ENV PORT=8000
ENV AWS_LWA_INVOKE_MODE=BUFFERED
ENV AWS_LWA_READINESS_CHECK_PATH=/mcp
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# soffice (LibreOffice) + poppler-utils are needed by the
# render_deck_to_images tool which converts the working presentation
# to per-slide PNGs so the assistant can visually inspect its build
# before saving. libreoffice-core + libreoffice-impress is the
# minimum subset for pptx → pdf headless conversion (~280MB);
# poppler-utils provides pdftoppm for pdf → png-per-page (~20MB).
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice-core \
        libreoffice-impress \
        poppler-utils \
        fonts-dejavu \
        fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8000

CMD ["python", "ppt_mcp_server.py", "--transport", "http", "--port", "8000"]
