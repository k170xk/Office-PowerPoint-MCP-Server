#!/usr/bin/env python3
"""
HTTP server wrapper for Office PowerPoint MCP Server.
Provides OpenAI-compatible JSON-RPC endpoints and presentation serving.
"""

import os
import json
import asyncio
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, unquote
import sys

# Add the project root to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import the FastMCP app and necessary components
# Import the entire module to ensure all tools are registered
import ppt_mcp_server
from ppt_mcp_server import app, presentations, current_presentation_id
from presentation_manager import get_presentation_manager
from storage_adapter import get_storage_adapter

# Presentation storage directory
PRESENTATIONS_DIR = os.getenv('PRESENTATIONS_DIR', './presentations')
BASE_URL = os.getenv('BASE_URL', '')  # Will be set from Render service URL

# Ensure presentations directory exists
os.makedirs(PRESENTATIONS_DIR, exist_ok=True)


class MCPHTTPHandler(BaseHTTPRequestHandler):
    """HTTP handler for MCP JSON-RPC requests and presentation serving."""
    
    def do_OPTIONS(self):
        """Handle CORS preflight requests."""
        self.send_response(200)
        self.send_cors_headers()
        self.end_headers()
    
    def send_cors_headers(self):
        """Send CORS headers."""
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.send_header('Content-Type', 'application/json')
    
    def do_GET(self):
        """Handle GET requests for tool discovery and presentation serving."""
        parsed_path = urlparse(self.path)
        path = parsed_path.path
        
        # Tool discovery endpoint
        if path == '/mcp/stream' or path == '/mcp/tools':
            try:
                request = {
                    "jsonrpc": "2.0",
                    "id": 1,
                    "method": "tools/list",
                    "params": {}
                }
                response = asyncio.run(self.handle_mcp_request(request))
                
                self.send_response(200)
                self.send_cors_headers()
                self.end_headers()
                self.wfile.write(json.dumps(response).encode('utf-8'))
            except Exception as e:
                self.send_error(500, f"Error: {str(e)}")
        
        # Presentation serving endpoint
        elif path.startswith('/presentations/'):
            filename = path.replace('/presentations/', '')
            self.serve_presentation(filename)
        
        # Health check
        elif path == '/health':
            self.send_response(200)
            self.send_cors_headers()
            self.end_headers()
            self.wfile.write(json.dumps({"status": "ok"}).encode('utf-8'))
        
        else:
            self.send_error(404, "Not found")
    
    def do_POST(self):
        """Handle POST requests for MCP tool calls."""
        parsed_path = urlparse(self.path)
        path = parsed_path.path
        
        if path == '/mcp/stream':
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length).decode('utf-8')
            
            try:
                request = json.loads(body)
                response = asyncio.run(self.handle_mcp_request(request))
                
                self.send_response(200)
                self.send_cors_headers()
                self.end_headers()
                self.wfile.write(json.dumps(response).encode('utf-8'))
            except json.JSONDecodeError as e:
                self.send_error(400, f"Invalid JSON: {str(e)}")
            except Exception as e:
                self.send_error(500, f"Error: {str(e)}")
        else:
            self.send_error(404, "Not found")
    
    async def handle_mcp_request(self, request: dict):
        """Handle an MCP JSON-RPC request using FastMCP."""
        method = request.get('method')
        params = request.get('params', {})
        request_id = request.get('id')
        
        # Try to use FastMCP's internal request handler first
        try:
            # FastMCP might have a handle_request method or similar
            if hasattr(app, 'handle_request') or hasattr(app, '_handle_request'):
                handler = getattr(app, 'handle_request', None) or getattr(app, '_handle_request', None)
                if handler:
                    result = await handler(request)
                    return result
        except Exception as e:
            print(f"FastMCP handler not available: {e}")
        
        try:
            if method == 'initialize':
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": {
                        "protocolVersion": "2024-11-05",
                        "capabilities": {
                            "tools": {}
                        },
                        "serverInfo": {
                            "name": "office-powerpoint-mcp-server",
                            "version": "2.0.0"
                        }
                    }
                }
            
            elif method == 'tools/list':
                # Get tools from FastMCP app
                tools = []
                # FastMCP stores tools internally - try multiple ways to access them
                try:
                    # Method 1: Try _tool_registry (common in FastMCP)
                    if hasattr(app, '_tool_registry'):
                        for tool_name, tool_info in app._tool_registry.items():
                            schema = tool_info.get('inputSchema', {}) if isinstance(tool_info, dict) else {}
                            desc = tool_info.get('description', f"Tool: {tool_name}") if isinstance(tool_info, dict) else (getattr(tool_info, '__doc__', None) or f"Tool: {tool_name}")
                            tools.append({
                                "name": tool_name,
                                "description": desc,
                                "inputSchema": schema if schema else {"type": "object", "properties": {}}
                            })
                    # Method 2: Try _tools attribute
                    elif hasattr(app, '_tools'):
                        for tool_name, tool_info in app._tools.items():
                            if isinstance(tool_info, dict):
                                tools.append({
                                    "name": tool_name,
                                    "description": tool_info.get('description', f"Tool: {tool_name}"),
                                    "inputSchema": tool_info.get('inputSchema', {"type": "object", "properties": {}})
                                })
                            else:
                                # tool_info might be a function
                                tools.append({
                                    "name": tool_name,
                                    "description": getattr(tool_info, '__doc__', None) or f"Tool: {tool_name}",
                                    "inputSchema": {"type": "object", "properties": {}}
                                })
                    # Method 3: Try accessing via __dict__ or vars()
                    elif hasattr(app, '__dict__'):
                        app_dict = vars(app)
                        for key, value in app_dict.items():
                            if 'tool' in key.lower() and isinstance(value, dict):
                                for tool_name, tool_info in value.items():
                                    tools.append({
                                        "name": tool_name,
                                        "description": tool_info.get('description', f"Tool: {tool_name}") if isinstance(tool_info, dict) else f"Tool: {tool_name}",
                                        "inputSchema": tool_info.get('inputSchema', {"type": "object", "properties": {}}) if isinstance(tool_info, dict) else {"type": "object", "properties": {}}
                                    })
                    # Method 4: Use FastMCP's internal server to handle the request
                    else:
                        # Create a mock request and use FastMCP's handler
                        from mcp.server.fastmcp import FastMCP
                        # Try to get tools via the server's internal handler
                        try:
                            # FastMCP might have a _server attribute
                            if hasattr(app, '_server') and hasattr(app._server, 'list_tools'):
                                result = await app._server.list_tools({})
                                if result and 'tools' in result:
                                    tools = result['tools']
                        except:
                            pass
                    
                    # If still no tools, try inspecting the app's registered functions
                    if not tools:
                        # Last resort: manually list known tools from the modules
                        known_tools = [
                            "create_presentation", "create_presentation_from_template", "open_presentation",
                            "save_presentation", "get_presentation_info", "get_template_file_info", "set_core_properties",
                            "add_slide", "get_slide_info", "extract_slide_text", "extract_presentation_text",
                            "populate_placeholder", "add_bullet_points", "manage_text", "manage_image",
                            "add_table", "format_table_cell", "add_shape", "add_chart", "update_chart_data",
                            "apply_professional_design", "apply_picture_effects", "manage_fonts",
                            "list_slide_templates", "apply_slide_template", "create_slide_from_template",
                            "create_presentation_from_templates", "get_template_info", "auto_generate_presentation",
                            "optimize_slide_text", "manage_hyperlinks", "add_connector",
                            "manage_slide_masters", "manage_slide_transitions",
                            "list_presentations", "switch_presentation", "get_server_info"
                        ]
                        for tool_name in known_tools:
                            tools.append({
                                "name": tool_name,
                                "description": f"Tool: {tool_name}",
                                "inputSchema": {"type": "object", "properties": {}}
                            })
                    
                except Exception as e:
                    import traceback
                    print(f"Warning: Could not access FastMCP tools: {e}")
                    traceback.print_exc()
                    # Fallback to known tools list
                    known_tools = [
                        "create_presentation", "create_presentation_from_template", "open_presentation",
                        "save_presentation", "get_presentation_info", "get_template_file_info", "set_core_properties",
                        "add_slide", "get_slide_info", "extract_slide_text", "extract_presentation_text",
                        "populate_placeholder", "add_bullet_points", "manage_text", "manage_image",
                        "add_table", "format_table_cell", "add_shape", "add_chart", "update_chart_data",
                        "apply_professional_design", "apply_picture_effects", "manage_fonts",
                        "list_slide_templates", "apply_slide_template", "create_slide_from_template",
                        "create_presentation_from_templates", "get_template_info", "auto_generate_presentation",
                        "optimize_slide_text", "manage_hyperlinks", "add_connector",
                        "manage_slide_masters", "manage_slide_transitions",
                        "list_presentations", "switch_presentation", "get_server_info"
                    ]
                    for tool_name in known_tools:
                        tools.append({
                            "name": tool_name,
                            "description": f"Tool: {tool_name}",
                            "inputSchema": {"type": "object", "properties": {}}
                        })
                
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": {
                        "tools": tools
                    }
                }
            
            elif method == 'tools/call':
                tool_name = params.get('name')
                arguments = params.get('arguments', {})
                
                # Handle file_path parameters - use storage adapter
                manager = get_presentation_manager()
                storage = get_storage_adapter()
                
                original_file_path = None
                local_path = None
                
                if 'file_path' in arguments:
                    original_file_path = arguments['file_path']
                    # Extract just the filename (remove path if present)
                    filename_base = os.path.basename(original_file_path)
                    
                    # Ensure .pptx extension
                    if not filename_base.endswith('.pptx'):
                        filename_base = f"{filename_base}.pptx"
                    
                    # Check if presentation exists in storage
                    create_if_missing = 'create' in tool_name or 'add' in tool_name
                    try:
                        local_path = manager.get_local_path(filename_base, create_if_missing=create_if_missing)
                        arguments['file_path'] = local_path
                    except FileNotFoundError:
                        if create_if_missing:
                            local_path = manager.get_local_path(filename_base, create_if_missing=True)
                            arguments['file_path'] = local_path
                        else:
                            return {
                                "jsonrpc": "2.0",
                                "id": request_id,
                                "error": {
                                    "code": -32602,
                                    "message": f"Presentation {filename_base} not found"
                                }
                            }
                
                # Call the tool via FastMCP's call_tool method
                try:
                    # Use FastMCP's internal tool calling mechanism
                    result = await app.call_tool(tool_name, arguments)
                    
                    # Upload presentation back to storage if it was modified
                    if local_path and os.path.exists(local_path):
                        if original_file_path:
                            filename_base = os.path.basename(original_file_path)
                        else:
                            filename_base = os.path.basename(arguments.get('file_path', ''))
                        
                        # Ensure .pptx extension
                        if filename_base and not filename_base.endswith('.pptx'):
                            filename_base = f"{filename_base}.pptx"
                        
                        if filename_base:
                            # Save to storage
                            pres_url = manager.save_presentation(local_path, filename_base)
                            # Enhance result with URL
                            if isinstance(result, dict):
                                result['download_url'] = f"{BASE_URL or 'https://office-powerpoint-mcp.onrender.com'}/presentations/{filename_base}"
                                if 'message' in result:
                                    result['message'] = f"{result['message']}\n\nPresentation saved: {filename_base}\nDownload URL: {result['download_url']}"
                            elif isinstance(result, str):
                                from urllib.parse import quote
                                encoded_filename = quote(filename_base)
                                download_url = f"{BASE_URL or 'https://office-powerpoint-mcp.onrender.com'}/presentations/{encoded_filename}"
                                result = f"{result}\n\nPresentation saved: {filename_base}\nDownload URL: {download_url}"
                    
                    # Convert result to string for JSON-RPC response
                    if isinstance(result, dict):
                        result_text = json.dumps(result, indent=2)
                    else:
                        result_text = str(result)
                    
                    enhanced_result = result_text
                except AttributeError:
                    # FastMCP might not have call_tool method, try direct access
                    try:
                        # Access tool function directly
                        tool_func = None
                        if hasattr(app, '_tools') and tool_name in app._tools:
                            tool_info = app._tools[tool_name]
                            tool_func = tool_info.get('handler') or tool_info.get('function')
                        elif hasattr(app, 'tools') and tool_name in app.tools:
                            tool_info = app.tools[tool_name]
                            tool_func = tool_info.get('handler') or tool_info.get('function')
                        
                        if tool_func is None:
                            return {
                                "jsonrpc": "2.0",
                                "id": request_id,
                                "error": {
                                    "code": -32601,
                                    "message": f"Tool not found: {tool_name}"
                                }
                            }
                        
                        # Call the tool function
                        if asyncio.iscoroutinefunction(tool_func):
                            result = await tool_func(**arguments)
                        else:
                            result = tool_func(**arguments)
                        
                        # Upload presentation back to storage if it was modified
                        if local_path and os.path.exists(local_path):
                            if original_file_path:
                                filename_base = os.path.basename(original_file_path)
                            else:
                                filename_base = os.path.basename(arguments.get('file_path', ''))
                            
                            # Ensure .pptx extension
                            if filename_base and not filename_base.endswith('.pptx'):
                                filename_base = f"{filename_base}.pptx"
                            
                            if filename_base:
                                # Save to storage
                                pres_url = manager.save_presentation(local_path, filename_base)
                                # Enhance result with URL
                                if isinstance(result, dict):
                                    result['download_url'] = f"{BASE_URL or 'https://office-powerpoint-mcp.onrender.com'}/presentations/{filename_base}"
                                    if 'message' in result:
                                        result['message'] = f"{result['message']}\n\nPresentation saved: {filename_base}\nDownload URL: {result['download_url']}"
                                elif isinstance(result, str):
                                    from urllib.parse import quote
                                    encoded_filename = quote(filename_base)
                                    download_url = f"{BASE_URL or 'https://office-powerpoint-mcp.onrender.com'}/presentations/{encoded_filename}"
                                    result = f"{result}\n\nPresentation saved: {filename_base}\nDownload URL: {download_url}"
                        
                        # Convert result to string for JSON-RPC response
                        if isinstance(result, dict):
                            result_text = json.dumps(result, indent=2)
                        else:
                            result_text = str(result)
                        
                        enhanced_result = result_text
                    except Exception as e:
                        return {
                            "jsonrpc": "2.0",
                            "id": request_id,
                            "error": {
                                "code": -32603,
                                "message": f"Error calling tool {tool_name}: {str(e)}"
                            }
                        }
                finally:
                    # Cleanup temp files
                    if local_path and os.path.exists(local_path):
                        manager.cleanup_temp(os.path.basename(local_path))
                
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": {
                        "content": [
                            {
                                "type": "text",
                                "text": enhanced_result
                            }
                        ]
                    }
                }
            
            else:
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "error": {
                        "code": -32601,
                        "message": f"Method not found: {method}"
                    }
                }
        
        except Exception as e:
            import traceback
            traceback.print_exc()
            return {
                "jsonrpc": "2.0",
                "id": request_id,
                "error": {
                    "code": -32603,
                    "message": str(e)
                }
            }
    
    def serve_presentation(self, filename: str):
        """Serve a presentation file from storage."""
        # URL decode the filename (handle %20 for spaces, etc.)
        filename = unquote(filename)
        # Security: prevent directory traversal
        filename = os.path.basename(filename)
        
        # Ensure .pptx extension
        if not filename.endswith('.pptx'):
            filename = f"{filename}.pptx"
        
        try:
            storage = get_storage_adapter()
            manager = get_presentation_manager()
            
            # Check if presentation exists in storage first
            if not storage.presentation_exists(filename):
                self.send_error(404, f"Presentation '{filename}' not found")
                return
            
            # Download from storage to temp location
            local_path = manager.get_local_path(filename, create_if_missing=False)
            
            if not os.path.exists(local_path):
                self.send_error(404, f"Presentation '{filename}' not found on disk")
                return
            
            with open(local_path, 'rb') as f:
                content = f.read()
            
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
            self.send_header('Content-Length', str(len(content)))
            self.send_cors_headers()
            self.end_headers()
            self.wfile.write(content)
            
            # Cleanup temp file
            manager.cleanup_temp(filename)
        except FileNotFoundError as e:
            self.send_error(404, f"Presentation not found: {str(e)}")
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.send_error(500, f"Error serving presentation: {str(e)}")
    
    def log_message(self, format, *args):
        """Override to use print instead of stderr."""
        print(f"{self.address_string()} - {format % args}")


def run_http_server():
    """Run the HTTP server."""
    port = int(os.getenv('PORT', 8000))
    host = os.getenv('HOST', '0.0.0.0')
    
    # Set BASE_URL if not already set
    global BASE_URL
    if not BASE_URL:
        # Try to get from Render environment
        render_service_url = os.getenv('RENDER_SERVICE_URL')
        if render_service_url:
            BASE_URL = render_service_url
        else:
            BASE_URL = f"http://{host}:{port}"
    
    server = HTTPServer((host, port), MCPHTTPHandler)
    print(f"Office PowerPoint MCP Server running on http://{host}:{port}")
    print(f"Presentations directory: {PRESENTATIONS_DIR}")
    print(f"Base URL: {BASE_URL}")
    print(f"MCP endpoint: http://{host}:{port}/mcp/stream")
    print(f"Presentations endpoint: http://{host}:{port}/presentations/")
    
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down server...")
        server.shutdown()


if __name__ == "__main__":
    run_http_server()

