#!/bin/bash

# API Testing Script for Student Marksheet Parser

echo "======================================"
echo "Student Marksheet Parser - API Tests"
echo "======================================"

API_URL="http://localhost:5000"

# Colors for output
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo ""
echo "Test 1: Health Check"
echo "------------------------------------"
response=$(curl -s -w "\n%{http_code}" "$API_URL/health")
http_code=$(echo "$response" | tail -n1)
body=$(echo "$response" | head -n-1)

if [ "$http_code" == "200" ]; then
    echo -e "${GREEN}✓ Health check passed${NC}"
    echo "$body" | python3 -m json.tool 2>/dev/null || echo "$body"
else
    echo -e "${RED}✗ Health check failed (HTTP $http_code)${NC}"
fi

echo ""
echo "Test 2: Upload Marksheet"
echo "------------------------------------"

# Check if sample file exists
if [ ! -f "sample_marksheet.xlsx" ]; then
    echo -e "${RED}✗ sample_marksheet.xlsx not found!${NC}"
    echo "Please create a sample marksheet file first."
    exit 1
fi

echo "Uploading sample_marksheet.xlsx..."
response=$(curl -s -w "\n%{http_code}" -X POST -F "file=@sample_marksheet.xlsx" "$API_URL/upload")
http_code=$(echo "$response" | tail -n1)
body=$(echo "$response" | head -n-1)

if [ "$http_code" == "200" ]; then
    echo -e "${GREEN}✓ Upload successful${NC}"
    echo "$body" | python3 -m json.tool 2>/dev/null || echo "$body"
    
    # Extract download URL
    download_url=$(echo "$body" | python3 -c "import sys, json; print(json.load(sys.stdin)['download_url'])" 2>/dev/null)
    
    if [ ! -z "$download_url" ]; then
        echo ""
        echo "Test 3: Download Reports"
        echo "------------------------------------"
        filename=$(basename "$download_url")
        echo "Downloading $filename..."
        
        curl -s -o "$filename" "$API_URL$download_url"
        
        if [ -f "$filename" ]; then
            size=$(ls -lh "$filename" | awk '{print $5}')
            echo -e "${GREEN}✓ Download successful${NC}"
            echo "File: $filename (Size: $size)"
        else
            echo -e "${RED}✗ Download failed${NC}"
        fi
    fi
else
    echo -e "${RED}✗ Upload failed (HTTP $http_code)${NC}"
    echo "$body"
fi

echo ""
echo "======================================"
echo "Testing Complete"
echo "======================================"