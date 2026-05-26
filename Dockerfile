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

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8000

CMD ["python", "ppt_mcp_server.py", "--transport", "http", "--port", "8000"]
