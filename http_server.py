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
import inspect
import typing

# Add the project root to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import the FastMCP app and necessary components
# Import the entire module to ensure all tools are registered
import ppt_mcp_server
from ppt_mcp_server import app, presentations, current_presentation_id
from presentation_manager import get_presentation_manager
from storage_adapter import get_storage_adapter

# Track presentation_id to filename mapping for auto-save
_presentation_files = {}  # Maps presentation_id to filename

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
                # Known tools list (fallback)
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
                
                # FastMCP stores tools internally - try multiple ways to access them
                try:
                    # Method 1: Try _tool_registry (common in FastMCP)
                    if hasattr(app, '_tool_registry'):
                        for tool_name, tool_info in app._tool_registry.items():
                            tool_func = tool_info if callable(tool_info) else (tool_info.get('handler') if isinstance(tool_info, dict) else None)
                            if tool_func:
                                schema = self._get_tool_schema(tool_func)
                            else:
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
                                tool_func = tool_info.get('handler') or tool_info.get('function')
                                if tool_func and callable(tool_func):
                                    schema = self._get_tool_schema(tool_func)
                                else:
                                    schema = tool_info.get('inputSchema', {})
                                tools.append({
                                    "name": tool_name,
                                    "description": tool_info.get('description', f"Tool: {tool_name}"),
                                    "inputSchema": schema if schema else {"type": "object", "properties": {}}
                                })
                            else:
                                # tool_info might be a function
                                if callable(tool_info):
                                    schema = self._get_tool_schema(tool_info)
                                else:
                                    schema = {"type": "object", "properties": {}}
                                tools.append({
                                    "name": tool_name,
                                    "description": getattr(tool_info, '__doc__', None) or f"Tool: {tool_name}",
                                    "inputSchema": schema
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
                    
                    # If still no tools, use fallback - try to get functions from ppt_mcp_server module
                    if not tools:
                        print("Warning: Could not access FastMCP tools, trying to extract from module")
                        # Try to get tool functions from the registered modules
                        try:
                            # Import tool modules to access functions
                            from tools import presentation_tools, content_tools, structural_tools, professional_tools
                            from tools import template_tools, hyperlink_tools, chart_tools, connector_tools
                            from tools import master_tools, transition_tools
                            
                            # Map tool names to their modules/functions
                            tool_modules = {
                                'create_presentation': (presentation_tools, 'create_presentation'),
                                'create_presentation_from_template': (presentation_tools, 'create_presentation_from_template'),
                                'open_presentation': (presentation_tools, 'open_presentation'),
                                'save_presentation': (presentation_tools, 'save_presentation'),
                                'get_presentation_info': (presentation_tools, 'get_presentation_info'),
                                'get_template_file_info': (presentation_tools, 'get_template_file_info'),
                                'set_core_properties': (presentation_tools, 'set_core_properties'),
                                'add_slide': (content_tools, 'add_slide'),
                                'get_slide_info': (content_tools, 'get_slide_info'),
                                'extract_slide_text': (content_tools, 'extract_slide_text'),
                                'extract_presentation_text': (content_tools, 'extract_presentation_text'),
                                'populate_placeholder': (content_tools, 'populate_placeholder'),
                                'add_bullet_points': (content_tools, 'add_bullet_points'),
                                'manage_text': (content_tools, 'manage_text'),
                                'manage_image': (content_tools, 'manage_image'),
                                'add_table': (structural_tools, 'add_table'),
                                'format_table_cell': (structural_tools, 'format_table_cell'),
                                'add_shape': (structural_tools, 'add_shape'),
                                'add_chart': (structural_tools, 'add_chart'),
                                'update_chart_data': (chart_tools, 'update_chart_data'),
                                'apply_professional_design': (professional_tools, 'apply_professional_design'),
                                'apply_picture_effects': (professional_tools, 'apply_picture_effects'),
                                'manage_fonts': (professional_tools, 'manage_fonts'),
                                'list_slide_templates': (template_tools, 'list_slide_templates'),
                                'apply_slide_template': (template_tools, 'apply_slide_template'),
                                'create_slide_from_template': (template_tools, 'create_slide_from_template'),
                                'create_presentation_from_templates': (template_tools, 'create_presentation_from_templates'),
                                'get_template_info': (template_tools, 'get_template_info'),
                                'auto_generate_presentation': (template_tools, 'auto_generate_presentation'),
                                'optimize_slide_text': (template_tools, 'optimize_slide_text'),
                                'manage_hyperlinks': (hyperlink_tools, 'manage_hyperlinks'),
                                'add_connector': (connector_tools, 'add_connector'),
                                'manage_slide_masters': (master_tools, 'manage_slide_masters'),
                                'manage_slide_transitions': (transition_tools, 'manage_slide_transitions'),
                            }
                            
                            # Also check ppt_mcp_server for additional tools
                            tool_modules['list_presentations'] = (ppt_mcp_server, 'list_presentations')
                            tool_modules['switch_presentation'] = (ppt_mcp_server, 'switch_presentation')
                            tool_modules['get_server_info'] = (ppt_mcp_server, 'get_server_info')
                            
                            for tool_name in known_tools:
                                if tool_name in tool_modules:
                                    module, func_name = tool_modules[tool_name]
                                    try:
                                        tool_func = getattr(module, func_name, None)
                                        if tool_func and callable(tool_func):
                                            schema = self._get_tool_schema(tool_func)
                                            desc = getattr(tool_func, '__doc__', None) or f"Tool: {tool_name}"
                                            tools.append({
                                                "name": tool_name,
                                                "description": desc.strip() if desc else f"Tool: {tool_name}",
                                                "inputSchema": schema
                                            })
                                            continue
                                    except Exception as e:
                                        print(f"Error extracting schema for {tool_name}: {e}")
                                
                                # Fallback if function not found
                                tools.append({
                                    "name": tool_name,
                                    "description": f"Tool: {tool_name}",
                                    "inputSchema": {"type": "object", "properties": {}}
                                })
                        except Exception as e:
                            print(f"Error extracting tools from modules: {e}")
                            # Final fallback
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
                    # Fallback - try to extract from modules (same logic as above)
                    try:
                        from tools import presentation_tools, content_tools, structural_tools, professional_tools
                        from tools import template_tools, hyperlink_tools, chart_tools, connector_tools
                        from tools import master_tools, transition_tools
                        
                        tool_modules = {
                            'create_presentation': (presentation_tools, 'create_presentation'),
                            'create_presentation_from_template': (presentation_tools, 'create_presentation_from_template'),
                            'open_presentation': (presentation_tools, 'open_presentation'),
                            'save_presentation': (presentation_tools, 'save_presentation'),
                            'get_presentation_info': (presentation_tools, 'get_presentation_info'),
                            'get_template_file_info': (presentation_tools, 'get_template_file_info'),
                            'set_core_properties': (presentation_tools, 'set_core_properties'),
                            'add_slide': (content_tools, 'add_slide'),
                            'get_slide_info': (content_tools, 'get_slide_info'),
                            'extract_slide_text': (content_tools, 'extract_slide_text'),
                            'extract_presentation_text': (content_tools, 'extract_presentation_text'),
                            'populate_placeholder': (content_tools, 'populate_placeholder'),
                            'add_bullet_points': (content_tools, 'add_bullet_points'),
                            'manage_text': (content_tools, 'manage_text'),
                            'manage_image': (content_tools, 'manage_image'),
                            'add_table': (structural_tools, 'add_table'),
                            'format_table_cell': (structural_tools, 'format_table_cell'),
                            'add_shape': (structural_tools, 'add_shape'),
                            'add_chart': (structural_tools, 'add_chart'),
                            'update_chart_data': (chart_tools, 'update_chart_data'),
                            'apply_professional_design': (professional_tools, 'apply_professional_design'),
                            'apply_picture_effects': (professional_tools, 'apply_picture_effects'),
                            'manage_fonts': (professional_tools, 'manage_fonts'),
                            'list_slide_templates': (template_tools, 'list_slide_templates'),
                            'apply_slide_template': (template_tools, 'apply_slide_template'),
                            'create_slide_from_template': (template_tools, 'create_slide_from_template'),
                            'create_presentation_from_templates': (template_tools, 'create_presentation_from_templates'),
                            'get_template_info': (template_tools, 'get_template_info'),
                            'auto_generate_presentation': (template_tools, 'auto_generate_presentation'),
                            'optimize_slide_text': (template_tools, 'optimize_slide_text'),
                            'manage_hyperlinks': (hyperlink_tools, 'manage_hyperlinks'),
                            'add_connector': (connector_tools, 'add_connector'),
                            'manage_slide_masters': (master_tools, 'manage_slide_masters'),
                            'manage_slide_transitions': (transition_tools, 'manage_slide_transitions'),
                            'list_presentations': (ppt_mcp_server, 'list_presentations'),
                            'switch_presentation': (ppt_mcp_server, 'switch_presentation'),
                            'get_server_info': (ppt_mcp_server, 'get_server_info'),
                        }
                        
                        for tool_name in known_tools:
                            if tool_name in tool_modules:
                                module, func_name = tool_modules[tool_name]
                                try:
                                    tool_func = getattr(module, func_name, None)
                                    if tool_func and callable(tool_func):
                                        schema = self._get_tool_schema(tool_func)
                                        desc = getattr(tool_func, '__doc__', None) or f"Tool: {tool_name}"
                                        tools.append({
                                            "name": tool_name,
                                            "description": desc.strip() if desc else f"Tool: {tool_name}",
                                            "inputSchema": schema
                                        })
                                        continue
                                except:
                                    pass
                            
                            tools.append({
                                "name": tool_name,
                                "description": f"Tool: {tool_name}",
                                "inputSchema": {"type": "object", "properties": {}}
                            })
                    except:
                        # Final fallback
                        for tool_name in known_tools:
                            tools.append({
                                "name": tool_name,
                                "description": f"Tool: {tool_name}",
                                "inputSchema": {"type": "object", "properties": {}}
                            })
                
                # Ensure we always return at least the known tools
                if not tools:
                    print("Warning: No tools found, using fallback list")
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
                filename_base = None
                
                # Handle file_path parameter (for open_presentation, save_presentation, create_presentation_from_template)
                if 'file_path' in arguments:
                    original_file_path = arguments['file_path']
                    # Extract just the filename (remove path if present)
                    filename_base = os.path.basename(original_file_path)
                    
                    # Ensure .pptx extension
                    if not filename_base.endswith('.pptx'):
                        filename_base = f"{filename_base}.pptx"
                    
                    # Check if presentation exists in storage
                    create_if_missing = 'create' in tool_name or 'add' in tool_name or 'from_template' in tool_name
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
                
                # Handle template_path parameter (for template operations)
                if 'template_path' in arguments:
                    template_path = arguments['template_path']
                    # Check if template exists in storage or templates directory
                    template_filename = os.path.basename(template_path)
                    if not template_filename.endswith('.pptx'):
                        template_filename = f"{template_filename}.pptx"
                    
                    # Try to find template in storage first
                    if storage.presentation_exists(template_filename):
                        local_template_path = manager.get_local_path(template_filename, create_if_missing=False)
                        arguments['template_path'] = local_template_path
                    else:
                        # Check in templates directory or PPT_TEMPLATE_PATH
                        template_dirs = [
                            os.getenv('PPT_TEMPLATE_PATH', ''),
                            '/mnt/disk/presentations/templates',
                            '/mnt/disk/templates',
                            './templates',
                            '.'
                        ]
                        template_found = False
                        for template_dir in template_dirs:
                            if template_dir and os.path.exists(template_dir):
                                potential_path = os.path.join(template_dir, template_filename)
                                if os.path.exists(potential_path):
                                    arguments['template_path'] = potential_path
                                    template_found = True
                                    break
                                # Also try with original name
                                potential_path = os.path.join(template_dir, template_path)
                                if os.path.exists(potential_path):
                                    arguments['template_path'] = potential_path
                                    template_found = True
                                    break
                        
                        if not template_found and not os.path.exists(template_path):
                            # Template not found - will be handled by the tool itself
                            pass
                
                # Call the tool via FastMCP's call_tool method
                try:
                    # Use FastMCP's internal tool calling mechanism
                    result = await app.call_tool(tool_name, arguments)
                    
                    # Track presentation_id to filename mapping
                    if isinstance(result, dict):
                        pres_id = result.get('presentation_id')
                        if pres_id and filename_base:
                            _presentation_files[pres_id] = filename_base
                    
                    # Handle save_presentation - upload to storage
                    if tool_name == 'save_presentation' and local_path and os.path.exists(local_path):
                        if original_file_path:
                            filename_base = os.path.basename(original_file_path)
                        else:
                            filename_base = os.path.basename(arguments.get('file_path', ''))
                        
                        if filename_base and not filename_base.endswith('.pptx'):
                            filename_base = f"{filename_base}.pptx"
                        
                        if filename_base:
                            pres_url = manager.save_presentation(local_path, filename_base)
                            if isinstance(result, dict) and 'presentation_id' in result:
                                _presentation_files[result['presentation_id']] = filename_base
                    
                    # Auto-save after modifications
                    modification_tools = [
                        'add_slide', 'manage_text', 'manage_image', 'add_table', 'format_table_cell',
                        'add_shape', 'add_chart', 'update_chart_data', 'apply_professional_design',
                        'apply_picture_effects', 'manage_fonts', 'apply_slide_template',
                        'create_slide_from_template', 'populate_placeholder', 'add_bullet_points',
                        'manage_hyperlinks', 'add_connector', 'manage_slide_masters',
                        'manage_slide_transitions', 'optimize_slide_text'
                    ]
                    
                    if tool_name in modification_tools:
                        pres_id = arguments.get('presentation_id') or current_presentation_id
                        if pres_id and pres_id in _presentation_files:
                            auto_save_filename = _presentation_files[pres_id]
                            if pres_id in presentations:
                                import tempfile
                                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
                                temp_path = temp_file.name
                                temp_file.close()
                                try:
                                    from utils import presentation_utils as ppt_utils
                                    ppt_utils.save_presentation(presentations[pres_id], temp_path)
                                    if os.path.exists(temp_path):
                                        manager.save_presentation(temp_path, auto_save_filename)
                                        os.unlink(temp_path)
                                except Exception as e:
                                    print(f"Warning: Auto-save failed: {e}")
                                    if os.path.exists(temp_path):
                                        os.unlink(temp_path)
                    
                    # Upload presentation back to storage if file_path was used
                    if local_path and os.path.exists(local_path) and tool_name != 'save_presentation':
                        if original_file_path:
                            filename_base = os.path.basename(original_file_path)
                        else:
                            filename_base = os.path.basename(arguments.get('file_path', ''))
                        
                        if filename_base and not filename_base.endswith('.pptx'):
                            filename_base = f"{filename_base}.pptx"
                        
                        if filename_base:
                            pres_url = manager.save_presentation(local_path, filename_base)
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
                        
                        # Track presentation_id to filename mapping for auto-save
                        if isinstance(result, dict):
                            pres_id = result.get('presentation_id')
                            if pres_id and filename_base:
                                _presentation_files[pres_id] = filename_base
                        
                        # Handle save_presentation - upload to storage
                        if tool_name == 'save_presentation' and local_path and os.path.exists(local_path):
                            if original_file_path:
                                filename_base = os.path.basename(original_file_path)
                            else:
                                filename_base = os.path.basename(arguments.get('file_path', ''))
                            
                            if filename_base and not filename_base.endswith('.pptx'):
                                filename_base = f"{filename_base}.pptx"
                            
                            if filename_base:
                                # Save to storage
                                pres_url = manager.save_presentation(local_path, filename_base)
                                # Update mapping if presentation_id is in result
                                if isinstance(result, dict) and 'presentation_id' in result:
                                    _presentation_files[result['presentation_id']] = filename_base
                        
                        # Auto-save presentations after modifications (for tools that modify presentations)
                        modification_tools = [
                            'add_slide', 'manage_text', 'manage_image', 'add_table', 'format_table_cell',
                            'add_shape', 'add_chart', 'update_chart_data', 'apply_professional_design',
                            'apply_picture_effects', 'manage_fonts', 'apply_slide_template',
                            'create_slide_from_template', 'populate_placeholder', 'add_bullet_points',
                            'manage_hyperlinks', 'add_connector', 'manage_slide_masters',
                            'manage_slide_transitions', 'optimize_slide_text'
                        ]
                        
                        if tool_name in modification_tools:
                            # Get presentation_id from arguments or current
                            pres_id = arguments.get('presentation_id') or current_presentation_id
                            if pres_id and pres_id in _presentation_files:
                                # Get the filename for this presentation
                                auto_save_filename = _presentation_files[pres_id]
                                if pres_id in presentations:
                                    # Save the in-memory presentation to a temp file, then upload
                                    import tempfile
                                    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
                                    temp_path = temp_file.name
                                    temp_file.close()
                                    
                                    try:
                                        from utils import presentation_utils as ppt_utils
                                        ppt_utils.save_presentation(presentations[pres_id], temp_path)
                                        if os.path.exists(temp_path):
                                            # Upload to storage
                                            manager.save_presentation(temp_path, auto_save_filename)
                                            # Cleanup
                                            os.unlink(temp_path)
                                    except Exception as e:
                                        print(f"Warning: Auto-save failed for {pres_id}: {e}")
                                        if os.path.exists(temp_path):
                                            os.unlink(temp_path)
                        
                        # Upload presentation back to storage if file_path was used and file exists
                        if local_path and os.path.exists(local_path) and tool_name != 'save_presentation':
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
    
    def _get_tool_schema(self, tool_func):
        """Extract JSON schema from tool function signature."""
        sig = inspect.signature(tool_func)
        
        properties = {}
        required = []
        
        # Get docstring for better descriptions
        docstring = tool_func.__doc__ or ""
        
        for param_name, param in sig.parameters.items():
            if param_name == 'self':
                continue
            
            param_type = param.annotation
            param_default = param.default
            
            # Handle Optional types
            if hasattr(typing, 'get_origin') and typing.get_origin(param_type) is typing.Union:
                args = typing.get_args(param_type)
                # If Union includes None, it's Optional
                if type(None) in args:
                    # Get the actual type (not None)
                    param_type = next((arg for arg in args if arg is not type(None)), str)
            
            # Map Python types to JSON schema types
            prop_schema = {}
            
            if param_type == str or param_type == inspect.Parameter.empty or param_type == type(None):
                prop_schema["type"] = "string"
            elif param_type == int:
                prop_schema["type"] = "integer"
            elif param_type == float:
                prop_schema["type"] = "number"
            elif param_type == bool:
                prop_schema["type"] = "boolean"
            elif param_type == list or (hasattr(typing, '_GenericAlias') and 'list' in str(param_type)):
                prop_schema["type"] = "array"
                prop_schema["items"] = {"type": "string"}  # Default to string array
            elif param_type == dict:
                prop_schema["type"] = "object"
            else:
                prop_schema["type"] = "string"
            
            # Try to extract description from docstring
            desc = f"Parameter: {param_name}"
            if docstring:
                # Look for param_name in docstring
                import re
                pattern = rf"{param_name}:\s*([^\n]+)"
                match = re.search(pattern, docstring)
                if match:
                    desc = match.group(1).strip()
            prop_schema["description"] = desc
            
            properties[param_name] = prop_schema
            
            # Only require if no default value
            if param_default == inspect.Parameter.empty:
                required.append(param_name)
        
        schema = {
            "type": "object",
            "properties": properties
        }
        
        if required:
            schema["required"] = required
        
        return schema
    
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

