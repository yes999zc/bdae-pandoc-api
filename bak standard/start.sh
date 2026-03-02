#!/bin/bash
docker rm -f pandoc-api 2>/dev/null
docker run -d --name pandoc-api \
  --network docker_default \
  -p 5050:5050 \
  -v ~/AI_Workspace/07_Pandoc-api:/app/templates \
  pandoc-api
echo "pandoc-api started"
