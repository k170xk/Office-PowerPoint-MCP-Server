# Render Deployment Guide

This document explains how to deploy the Office PowerPoint MCP Server on Render.

## Required Environment Variables

Set the following environment variables in your Render service:

### `DISK_PATH`
- **Value**: `/mnt/disk/presentations` (or your custom path)
- **Description**: Path to Render Disk for persistent presentation storage
- **Required**: Yes (for persistent storage)

### `BASE_URL`
- **Value**: `https://your-service-name.onrender.com` (auto-set by Render)
- **Description**: Base URL for the service (used for presentation download URLs)
- **Required**: No (will be auto-detected from Render)

### `STORAGE_TYPE`
- **Value**: `disk` (default), `s3`, or `local`
- **Description**: Storage backend type. Use `disk` for Render Disk persistence
- **Required**: No (defaults to `disk`)

### `PORT`
- **Value**: Auto-set by Render (typically 10000)
- **Description**: Port number for the HTTP server
- **Required**: No (defaults to 8000, but Render sets this automatically)

### `PRESENTATIONS_DIR`
- **Value**: `./presentations` (fallback if disk not available)
- **Description**: Local directory for presentations (fallback only)
- **Required**: No

## How to Set Environment Variables

1. Go to your Render dashboard: https://dashboard.render.com
2. Navigate to your service: `Office-PowerPoint-MCP-Server`
3. Click on "Environment" in the left sidebar
4. Add the environment variables:
   - Key: `DISK_PATH`
   - Value: `/mnt/disk/presentations`
   - Key: `STORAGE_TYPE`
   - Value: `disk`
5. Click "Save Changes"

## Render Disk Setup

For persistent storage, attach a Render Disk to your service:

1. In your Render service settings, go to "Disks"
2. Click "Attach Disk"
3. Create a new disk or attach an existing one
4. Mount point: `/mnt/disk`
5. The server will automatically use `/mnt/disk/presentations` for storage

## Deployment

After setting the environment variables:

1. Render will automatically redeploy your service
2. The server will start with HTTP transport on the port provided by Render
3. Access your server at: `https://your-service-name.onrender.com/mcp/stream`

## Endpoints

- **MCP Stream**: `https://your-service-name.onrender.com/mcp/stream`
- **Tools List**: `https://your-service-name.onrender.com/mcp/tools`
- **Health Check**: `https://your-service-name.onrender.com/health`
- **Download Presentation**: `https://your-service-name.onrender.com/presentations/{filename}.pptx`

## Health Check Endpoint

The HTTP server provides a health check endpoint at:
- `https://your-service-name.onrender.com/health`

Configure this in Render's health check settings.

## Troubleshooting

### Server exits with status 1
- **Cause**: Missing dependencies or import errors
- **Fix**: Check Render logs for specific error messages

### Presentations not persisting
- **Cause**: Render Disk not attached or DISK_PATH not set correctly
- **Fix**: Ensure disk is attached at `/mnt/disk` and `DISK_PATH` is set to `/mnt/disk/presentations`

### Cannot connect to server
- **Cause**: Health checks failing or port binding issues
- **Fix**: Ensure `PORT` environment variable is set (Render sets this automatically)

### Tool calls failing
- **Cause**: FastMCP tool registry not accessible
- **Fix**: Check that `ppt_mcp_server.py` is properly imported and tools are registered

## Storage Options

### Render Disk (Recommended)
- Persistent across deployments
- No external service required
- Set `STORAGE_TYPE=disk` and `DISK_PATH=/mnt/disk/presentations`

### S3 (Alternative)
- Requires AWS credentials
- Set `STORAGE_TYPE=s3` and configure AWS environment variables:
  - `S3_BUCKET_NAME`
  - `AWS_ACCESS_KEY_ID`
  - `AWS_SECRET_ACCESS_KEY`
  - `S3_REGION`

### Local (Development Only)
- Ephemeral storage (lost on restart)
- Set `STORAGE_TYPE=local`
- Not recommended for production

