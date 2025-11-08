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
import ast
import importlib.util

# Add the project root to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import the FastMCP app and necessary components
# Import the entire module to ensure all tools are registered
try:
    import ppt_mcp_server
    from ppt_mcp_server import app, presentations, current_presentation_id
    from presentation_manager import get_presentation_manager
    from storage_adapter import get_storage_adapter
except ImportError as e:
    print(f"ERROR: Failed to import required modules: {e}")
    print("This usually means a dependency is missing. Check requirements.txt")
    import traceback
    traceback.print_exc()
    sys.exit(1)
except Exception as e:
    print(f"ERROR: Unexpected error during import: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Track presentation_id to filename mapping for auto-save
_presentation_files = {}  # Maps presentation_id to filename

# Build tool registry by extracting functions from FastMCP
TOOL_REGISTRY = {}

def build_tool_registry():
    """Build a registry of all available tools by extracting unwrapped functions from FastMCP."""
    global TOOL_REGISTRY
    
    # Try to use FastMCP's list_tools method first (it might have better schema extraction)
    try:
        import asyncio
        # FastMCP's list_tools might be async
        if hasattr(app, 'list_tools'):
            try:
                if asyncio.iscoroutinefunction(app.list_tools):
                    # Try to get existing event loop, or create new one
                    try:
                        loop = asyncio.get_event_loop()
                        if loop.is_running():
                            # Can't use asyncio.run() if loop is running, skip this
                            print("DEBUG: Event loop already running, skipping FastMCP list_tools")
                            tools_result = None
                        else:
                            tools_result = loop.run_until_complete(app.list_tools())
                    except RuntimeError:
                        # No event loop, create new one
                        tools_result = asyncio.run(app.list_tools())
                else:
                    tools_result = app.list_tools()
                
                if tools_result and 'tools' in tools_result:
                    fastmcp_tools = tools_result['tools']
                    # Check if schemas are populated
                    if fastmcp_tools and len(fastmcp_tools) > 0:
                        sample = fastmcp_tools[0]
                        if sample.get('inputSchema', {}).get('properties'):
                            print(f"✓ FastMCP list_tools provided {len(fastmcp_tools)} tools with schemas")
                            # Build registry from FastMCP's tool list
                            tools_map = {}
                            # We still need the actual functions for calling, so get them from app._tools
                            tools_dict = getattr(app, '_tools', None) or getattr(app, 'tools', None)
                            if tools_dict:
                                for tool_name, tool_info in tools_dict.items():
                                    tool_func = None
                                    if isinstance(tool_info, dict):
                                        tool_func = (tool_info.get('handler') or 
                                                    tool_info.get('function') or 
                                                    tool_info.get('func') or
                                                    tool_info.get('_func'))
                                    elif callable(tool_info):
                                        tool_func = tool_info
                                    
                                    if tool_func:
                                        # Unwrap function
                                        original_func = tool_func
                                        for _ in range(5):
                                            if hasattr(tool_func, '__wrapped__'):
                                                tool_func = tool_func.__wrapped__
                                            elif hasattr(tool_func, '_func'):
                                                tool_func = tool_func._func
                                            elif hasattr(tool_func, 'func'):
                                                tool_func = tool_func.func
                                            elif hasattr(tool_func, '__func__'):
                                                tool_func = tool_func.__func__
                                            else:
                                                break
                                        
                                        if not callable(tool_func):
                                            tool_func = original_func
                                        
                                        tools_map[tool_name] = tool_func
                            
                            TOOL_REGISTRY = tools_map
                            print(f"Built TOOL_REGISTRY with {len(TOOL_REGISTRY)} tools from FastMCP")
                            # Store the schemas for later use
                            return
            except Exception as e:
                print(f"DEBUG: FastMCP list_tools failed: {e}")
    except Exception as e:
        print(f"DEBUG: Error trying FastMCP list_tools: {e}")
    
    # Fallback: Access FastMCP's tools dict directly
    tools_dict = None
    if hasattr(app, '_tools'):
        tools_dict = app._tools
        print(f"DEBUG: Found app._tools with {len(tools_dict) if tools_dict else 0} items")
    elif hasattr(app, 'tools'):
        tools_dict = app.tools
        print(f"DEBUG: Found app.tools with {len(tools_dict) if tools_dict else 0} items")
    elif hasattr(app, '_tool_registry'):
        tools_dict = app._tool_registry
        print(f"DEBUG: Found app._tool_registry with {len(tools_dict) if tools_dict else 0} items")
    else:
        # Try to find tools in app.__dict__
        app_dict = vars(app) if hasattr(app, '__dict__') else {}
        for key, value in app_dict.items():
            if 'tool' in key.lower() and isinstance(value, dict):
                tools_dict = value
                print(f"DEBUG: Found tools in app.{key} with {len(tools_dict)} items")
                break
    
    if not tools_dict or len(tools_dict) == 0:
        print("Warning: FastMCP tools dict not found or empty")
        return
    
    # Extract unwrapped function objects
    tools_map = {}
    for tool_name, tool_info in tools_dict.items():
        tool_func = None
        
        # Get the actual function - FastMCP might wrap it in different ways
        if isinstance(tool_info, dict):
            # Try multiple keys to find the function
            tool_func = (tool_info.get('handler') or 
                        tool_info.get('function') or 
                        tool_info.get('func') or
                        tool_info.get('_func'))
        elif callable(tool_info):
            tool_func = tool_info
        
        if tool_func:
            # Unwrap if function is decorated (e.g., by FastMCP decorator)
            original_func = tool_func
            unwrap_attempts = 0
            max_unwrap = 5  # Prevent infinite loops
            
            while unwrap_attempts < max_unwrap:
                if hasattr(tool_func, '__wrapped__'):
                    tool_func = tool_func.__wrapped__
                    unwrap_attempts += 1
                elif hasattr(tool_func, '_func'):
                    tool_func = tool_func._func
                    unwrap_attempts += 1
                elif hasattr(tool_func, 'func'):
                    tool_func = tool_func.func
                    unwrap_attempts += 1
                elif hasattr(tool_func, '__func__'):
                    tool_func = tool_func.__func__
                    unwrap_attempts += 1
                else:
                    break
            
            # If we couldn't unwrap, use the original
            if not callable(tool_func):
                tool_func = original_func
            
            tools_map[tool_name] = tool_func
    
    TOOL_REGISTRY = tools_map
    print(f"Built TOOL_REGISTRY with {len(TOOL_REGISTRY)} tools")

# Build registry after tools are registered (ppt_mcp_server import registers all tools)
build_tool_registry()

# Presentation storage directory
PRESENTATIONS_DIR = os.getenv('PRESENTATIONS_DIR', './presentations')
BASE_URL = os.getenv('BASE_URL', '')  # Will be set from Render service URL

# Ensure BASE_URL is set for download URLs
if not BASE_URL:
    render_service_url = os.getenv('RENDER_SERVICE_URL')
    if render_service_url:
        BASE_URL = render_service_url
    else:
        # Default to the known service URL
        BASE_URL = 'https://office-powerpoint-mcp.onrender.com'

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
                # Try FastMCP's list_tools first (it might have better schema extraction)
                tools = []
                try:
                    if hasattr(app, 'list_tools'):
                        if asyncio.iscoroutinefunction(app.list_tools):
                            tools_result = await app.list_tools()
                        else:
                            tools_result = app.list_tools()
                        
                        if tools_result:
                            # Handle both dict response and direct tools list
                            if isinstance(tools_result, dict) and 'tools' in tools_result:
                                fastmcp_tools = tools_result['tools']
                            elif isinstance(tools_result, list):
                                fastmcp_tools = tools_result
                            else:
                                fastmcp_tools = []
                            
                            if fastmcp_tools and len(fastmcp_tools) > 0:
                                # Convert Tool objects to dicts if needed (FastMCP returns Pydantic models)
                                tools_list = []
                                for t in fastmcp_tools:
                                    if hasattr(t, 'model_dump'):
                                        # Pydantic v2
                                        tools_list.append(t.model_dump())
                                    elif hasattr(t, 'dict'):
                                        # Pydantic v1
                                        tools_list.append(t.dict())
                                    elif isinstance(t, dict):
                                        tools_list.append(t)
                                    else:
                                        # Try to convert manually
                                        tool_dict = {
                                            'name': getattr(t, 'name', ''),
                                            'description': getattr(t, 'description', ''),
                                            'inputSchema': getattr(t, 'inputSchema', {}) if hasattr(t, 'inputSchema') else {}
                                        }
                                        tools_list.append(tool_dict)
                                
                                # Check if schemas are populated - count how many have properties
                                tools_with_schemas = [t for t in tools_list if isinstance(t, dict) and t.get('inputSchema', {}).get('properties')]
                                if len(tools_with_schemas) >= len(tools_list) * 0.5:  # At least 50% have schemas
                                    print(f"✓ Using FastMCP list_tools with {len(tools_list)} tools ({len(tools_with_schemas)} with schemas)")
                                    tools = tools_list
                                else:
                                    print(f"FastMCP list_tools returned {len(tools_list)} tools but only {len(tools_with_schemas)} have schemas, will use TOOL_REGISTRY")
                except Exception as e:
                    print(f"FastMCP list_tools failed: {e}")
                    import traceback
                    traceback.print_exc()
                
                # Fallback: Use TOOL_REGISTRY with AST parsing for nested functions
                if not tools or not any(t.get('inputSchema', {}).get('properties') for t in tools):
                    print("Using TOOL_REGISTRY with AST parsing for schema extraction...")
                    tools = []
                    # Map tool names to their source modules (import inside try-except to handle missing modules)
                    tool_modules_map = {}  # Initialize to empty dict
                    try:
                        from tools import presentation_tools, content_tools, structural_tools, professional_tools
                        from tools import template_tools, hyperlink_tools, chart_tools, connector_tools
                        from tools import master_tools, transition_tools
                        
                        tool_modules_map = {
                            'create_presentation': presentation_tools,
                            'create_presentation_from_template': presentation_tools,
                            'open_presentation': presentation_tools,
                            'save_presentation': presentation_tools,
                            'get_presentation_info': presentation_tools,
                            'get_template_file_info': presentation_tools,
                            'set_core_properties': presentation_tools,
                            'add_slide': content_tools,
                            'get_slide_info': content_tools,
                            'extract_slide_text': content_tools,
                            'extract_presentation_text': content_tools,
                            'populate_placeholder': content_tools,
                            'add_bullet_points': content_tools,
                            'manage_text': content_tools,
                            'manage_image': content_tools,
                            'add_table': structural_tools,
                            'format_table_cell': structural_tools,
                            'add_shape': structural_tools,
                            'add_chart': structural_tools,
                            'update_chart_data': chart_tools,
                            'apply_professional_design': professional_tools,
                            'apply_picture_effects': professional_tools,
                            'manage_fonts': professional_tools,
                            'list_slide_templates': template_tools,
                            'apply_slide_template': template_tools,
                            'create_slide_from_template': template_tools,
                            'create_presentation_from_templates': template_tools,
                            'get_template_info': template_tools,
                            'auto_generate_presentation': template_tools,
                            'optimize_slide_text': template_tools,
                            'manage_hyperlinks': hyperlink_tools,
                            'add_connector': connector_tools,
                            'manage_slide_masters': master_tools,
                            'manage_slide_transitions': transition_tools,
                            'list_presentations': ppt_mcp_server,
                            'list_available_presentations': ppt_mcp_server,
                            'switch_presentation': ppt_mcp_server,
                            'get_server_info': ppt_mcp_server,
                        }
                    except ImportError as e:
                        print(f"Warning: Could not import tool modules: {e}")
                    except Exception as e:
                        print(f"Warning: Error setting up tool_modules_map: {e}")
                    
                    print(f"TOOL_REGISTRY has {len(TOOL_REGISTRY)} tools")
                    print(f"tool_modules_map has {len(tool_modules_map)} mappings")
                    
                    for tool_name, tool_func in TOOL_REGISTRY.items():
                        schema = None
                        # Try signature extraction first (more reliable than AST for wrapped functions)
                        try:
                            schema = self._get_tool_schema(tool_func)
                            if schema and schema.get('properties'):
                                print(f"  ✓ Extracted schema for {tool_name} from signature ({len(schema.get('properties', {}))} params)", flush=True)
                            else:
                                print(f"  ⚠ Signature extraction for {tool_name} returned empty schema", flush=True)
                        except Exception as e:
                            print(f"  ✗ Signature extraction failed for {tool_name}: {e}", flush=True)
                            import traceback
                            traceback.print_exc()
                        
                        # Fallback to AST parsing (works for nested functions)
                        if not schema or not schema.get('properties'):
                            if tool_name in tool_modules_map:
                                module = tool_modules_map[tool_name]
                                try:
                                    schema = self._get_tool_schema_from_source(module, tool_name, tool_func)
                                    if schema and schema.get('properties'):
                                        print(f"  ✓ Extracted schema for {tool_name} from source ({len(schema.get('properties', {}))} params)", flush=True)
                                    else:
                                        print(f"  ⚠ AST parsing for {tool_name} returned empty schema", flush=True)
                                except Exception as e:
                                    print(f"  ✗ AST parsing failed for {tool_name}: {e}", flush=True)
                                    import traceback
                                    traceback.print_exc()
                        
                        # Final fallback: empty schema
                        if not schema or not schema.get('properties'):
                            schema = {"type": "object", "properties": {}}
                            print(f"  ⚠ No schema extracted for {tool_name}, using empty schema", flush=True)
                        
                        tools.append({
                            "name": tool_name,
                            "description": tool_func.__doc__ or f"Tool: {tool_name}",
                            "inputSchema": schema
                        })
                    
                    # Check how many have schemas (including empty properties - valid for zero-param functions)
                    # A valid schema has type='object' - empty properties is valid for zero-parameter functions
                    schemas_count = sum(1 for t in tools if t.get('inputSchema', {}).get('type') == 'object')
                    tools_with_params = sum(1 for t in tools if t.get('inputSchema', {}).get('properties', {}))
                    print(f"TOOL_REGISTRY: {schemas_count}/{len(tools)} tools have schemas ({tools_with_params} with parameters, {schemas_count - tools_with_params} zero-param functions)", flush=True)
                    sys.stdout.flush()
                
                # Final fallback: if still empty, return known tools with empty schemas
                if not tools:
                    print("Warning: No tools found, using fallback list")
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
                        "list_presentations", "list_available_presentations", "switch_presentation", "get_server_info"
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
                
                # Use TOOL_REGISTRY (like Word MCP) - simple and reliable
                if tool_name not in TOOL_REGISTRY:
                    return {
                        "jsonrpc": "2.0",
                        "id": request_id,
                        "error": {
                            "code": -32601,
                            "message": f"Tool not found: {tool_name}"
                        }
                    }
                
                tool_func = TOOL_REGISTRY[tool_name]
                
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
                
                # Call the tool directly from TOOL_REGISTRY (like Word MCP)
                try:
                    # Call the function directly
                    if asyncio.iscoroutinefunction(tool_func):
                        result = await tool_func(**arguments)
                    else:
                        result = tool_func(**arguments)
                    
                    # Track presentation_id to filename mapping
                    if isinstance(result, dict):
                        pres_id = result.get('presentation_id')
                        if pres_id:
                            # For open_presentation, create_presentation_from_template, save_presentation
                            if filename_base:
                                _presentation_files[pres_id] = filename_base
                            # For create_presentation, if file_path was provided, use it
                            elif 'file_path' in arguments and arguments['file_path']:
                                # Extract filename from the file_path that was used
                                used_path = arguments['file_path']
                                if os.path.exists(used_path):
                                    # This is a local path, get the original filename
                                    if original_file_path:
                                        _presentation_files[pres_id] = os.path.basename(original_file_path)
                                    else:
                                        # Try to infer from local_path
                                        _presentation_files[pres_id] = os.path.basename(used_path)
                    
                    # Handle save_presentation - upload to storage
                    if tool_name == 'save_presentation' and local_path and os.path.exists(local_path):
                        if original_file_path:
                            filename_base = os.path.basename(original_file_path)
                        else:
                            filename_base = os.path.basename(arguments.get('file_path', ''))
                        
                        if filename_base and not filename_base.endswith('.pptx'):
                            filename_base = f"{filename_base}.pptx"
                        
                        if filename_base:
                            # Save to storage (this saves to Render Disk)
                            pres_url = manager.save_presentation(local_path, filename_base)
                            # Generate download URL
                            download_url = f"{BASE_URL or 'https://office-powerpoint-mcp.onrender.com'}/presentations/{filename_base}"
                            
                            # Verify file was saved
                            storage = get_storage_adapter()
                            if storage.presentation_exists(filename_base):
                                print(f"✓ Verified: Presentation {filename_base} exists in storage")
                            else:
                                print(f"⚠ Warning: Presentation {filename_base} not found in storage after save")
                            
                            # Enhance result with download URL
                            if isinstance(result, dict):
                                result['download_url'] = download_url
                                result['filename'] = filename_base
                                result['saved_to_disk'] = True
                                if 'presentation_id' in result:
                                    _presentation_files[result['presentation_id']] = filename_base
                                if 'message' in result:
                                    result['message'] = f"{result['message']}\n\n✓ Saved to Render Disk: {filename_base}\n📥 Download URL: {download_url}"
                                else:
                                    result['message'] = f"✓ Saved to Render Disk: {filename_base}\n📥 Download URL: {download_url}"
                            elif isinstance(result, str):
                                result = f"{result}\n\n✓ Saved to Render Disk: {filename_base}\n📥 Download URL: {download_url}"
                    
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
                                        # Save to storage and get URL (saves to Render Disk)
                                        pres_url = manager.save_presentation(temp_path, auto_save_filename)
                                        download_url = f"{BASE_URL or 'https://office-powerpoint-mcp.onrender.com'}/presentations/{auto_save_filename}"
                                        
                                        # Verify file was saved
                                        storage = get_storage_adapter()
                                        if storage.presentation_exists(auto_save_filename):
                                            print(f"✓ Verified: Auto-saved presentation {auto_save_filename} exists in storage")
                                        
                                        # Enhance result with download URL
                                        if isinstance(result, dict):
                                            if 'download_url' not in result:
                                                result['download_url'] = download_url
                                            result['filename'] = auto_save_filename
                                            result['auto_saved'] = True
                                            if 'message' in result:
                                                result['message'] = f"{result['message']}\n\n✓ Auto-saved to Render Disk: {auto_save_filename}\n📥 Download URL: {download_url}"
                                            else:
                                                result['message'] = f"✓ Auto-saved to Render Disk: {auto_save_filename}\n📥 Download URL: {download_url}"
                                        
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
                except Exception as e:
                    import traceback
                    traceback.print_exc()
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
    
    def _get_tool_schema_from_source(self, module, func_name, tool_func):
        """Extract schema by parsing the source file directly."""
        try:
            # Method 1: Try to get source directly from function object (works for nested functions)
            source_code = None
            if tool_func:
                try:
                    source_code = inspect.getsource(tool_func)
                    print(f"✓ Got source for {func_name} via inspect.getsource()")
                except (OSError, TypeError) as e:
                    print(f"  inspect.getsource() failed for {func_name}: {e}")
            
            # Method 2: Fallback to reading from file
            if not source_code:
                # Get source file from module (most reliable)
                source_file = None
                try:
                    # Try __file__ first (more reliable)
                    if hasattr(module, '__file__') and module.__file__:
                        source_file = module.__file__
                        # Handle .pyc files
                        if source_file.endswith('.pyc'):
                            source_file = source_file[:-1]  # Remove 'c' to get .py
                    else:
                        source_file = inspect.getfile(module)
                except (TypeError, OSError, AttributeError) as e:
                    print(f"Warning: Could not get source file from module {module.__name__}: {e}")
                    return {"type": "object", "properties": {}}
                
                # Verify file exists - try multiple paths
                if not os.path.exists(source_file):
                    # Try relative to current directory
                    basename = os.path.basename(source_file)
                    if os.path.exists(basename):
                        source_file = basename
                    # Try in tools/ directory
                    elif os.path.exists(f"tools/{basename}"):
                        source_file = f"tools/{basename}"
                    # Try absolute path from __file__
                    elif hasattr(module, '__file__') and module.__file__:
                        abs_path = os.path.abspath(module.__file__)
                        if abs_path.endswith('.pyc'):
                            abs_path = abs_path[:-1]
                        if os.path.exists(abs_path):
                            source_file = abs_path
                    else:
                        print(f"Warning: Source file not found: {source_file} (tried multiple paths)")
                        return {"type": "object", "properties": {}}
                
                # Read and parse the source file
                with open(source_file, 'r') as f:
                    source_code = f.read()
            
            # Parse AST
            tree = ast.parse(source_code)
            
            # Find the function definition (could be nested inside register function)
            # Try multiple strategies to find the function
            func_node = None
            
            # Strategy 1: Direct search by name
            for node in ast.walk(tree):
                if isinstance(node, ast.FunctionDef) and node.name == func_name:
                    func_node = node
                    break
            
            # Strategy 2: If not found, search in nested functions (closures)
            if not func_node:
                def find_function_recursive(node, target_name, depth=0):
                    """Recursively find function definition, even in nested closures."""
                    if isinstance(node, ast.FunctionDef):
                        if node.name == target_name:
                            return node
                        # Search inside this function's body
                        for child in node.body:
                            result = find_function_recursive(child, target_name, depth + 1)
                            if result:
                                return result
                    elif isinstance(node, (ast.AsyncFunctionDef, ast.Lambda)):
                        # Also check async functions and lambdas
                        if hasattr(node, 'name') and node.name == target_name:
                            return node
                    else:
                        # Search in all child nodes
                        for child in ast.iter_child_nodes(node):
                            result = find_function_recursive(child, target_name, depth + 1)
                            if result:
                                return result
                    return None
                
                func_node = find_function_recursive(tree, func_name)
            
            # Strategy 3: If still not found, try to match by docstring or other attributes
            if not func_node and tool_func:
                # Try to find by matching docstring
                try:
                    expected_doc = tool_func.__doc__ or ""
                    for node in ast.walk(tree):
                        if isinstance(node, ast.FunctionDef) and node.name == func_name:
                            # Check if docstring matches
                            if node.body:
                                first_stmt = node.body[0]
                                if isinstance(first_stmt, ast.Expr):
                                    if isinstance(first_stmt.value, ast.Str):
                                        docstring = first_stmt.value.s
                                    elif isinstance(first_stmt.value, ast.Constant) and isinstance(first_stmt.value.value, str):
                                        docstring = first_stmt.value.value
                                    else:
                                        docstring = ""
                                    if expected_doc and docstring and expected_doc.strip() in docstring:
                                        func_node = node
                                        break
                except:
                    pass
            if func_node:
                print(f"✓ Found function {func_name} in AST of {source_file}")
                # Found the function - extract parameters
                properties = {}
                required = []
                
                # Debug: print function args
                print(f"  Function has {len(func_node.args.args)} parameters")
                
                # Get all parameters
                args = func_node.args.args
                defaults = func_node.args.defaults
                num_defaults = len(defaults)
                num_args = len(args)
                
                for i, arg in enumerate(args):
                    if arg.arg == 'self':
                        continue
                    
                    param_name = arg.arg
                    param_schema = {}
                    
                    # Determine if parameter is required (no default value)
                    has_default = i >= (num_args - num_defaults)
                    if not has_default:
                        required.append(param_name)
                    
                    # Extract type annotation
                    if arg.annotation:
                        # Handle simple types
                        if isinstance(arg.annotation, ast.Name):
                            type_name = arg.annotation.id
                            if type_name == 'int':
                                param_schema["type"] = "integer"
                            elif type_name == 'float':
                                param_schema["type"] = "number"
                            elif type_name == 'bool':
                                param_schema["type"] = "boolean"
                            elif type_name == 'str':
                                param_schema["type"] = "string"
                            elif type_name == 'list' or type_name == 'List':
                                param_schema["type"] = "array"
                                param_schema["items"] = {"type": "string"}
                            elif type_name == 'dict' or type_name == 'Dict':
                                param_schema["type"] = "object"
                            else:
                                param_schema["type"] = "string"
                        
                        # Handle Optional[Type] - Python 3.8+ uses ast.Subscript with ast.Index
                        elif isinstance(arg.annotation, ast.Subscript):
                            if isinstance(arg.annotation.value, ast.Name) and arg.annotation.value.id == 'Optional':
                                # Get the inner type - handle both old and new AST formats
                                slice_value = arg.annotation.slice
                                if isinstance(slice_value, ast.Name):
                                    inner_type = slice_value.id
                                elif isinstance(slice_value, ast.Index):  # Python < 3.9
                                    if isinstance(slice_value.value, ast.Name):
                                        inner_type = slice_value.value.id
                                    else:
                                        inner_type = 'str'
                                else:
                                    inner_type = 'str'
                                
                                if inner_type == 'int':
                                    param_schema["type"] = "integer"
                                elif inner_type == 'float':
                                    param_schema["type"] = "number"
                                elif inner_type == 'bool':
                                    param_schema["type"] = "boolean"
                                elif inner_type == 'str':
                                    param_schema["type"] = "string"
                                elif inner_type == 'list' or inner_type == 'List':
                                    param_schema["type"] = "array"
                                    param_schema["items"] = {"type": "string"}
                                else:
                                    param_schema["type"] = "string"
                            else:
                                param_schema["type"] = "string"
                        else:
                            param_schema["type"] = "string"
                    else:
                        param_schema["type"] = "string"
                    
                    # Try to get description from docstring (handle both ast.Str and ast.Constant)
                    docstring = None
                    if func_node.body:
                        first_stmt = func_node.body[0]
                        if isinstance(first_stmt, ast.Expr):
                            if isinstance(first_stmt.value, ast.Str):  # Python < 3.8
                                docstring = first_stmt.value.s
                            elif isinstance(first_stmt.value, ast.Constant) and isinstance(first_stmt.value.value, str):  # Python 3.8+
                                docstring = first_stmt.value.value
                    
                    if docstring:
                        # Look for param_name in docstring
                        import re
                        pattern = rf"{param_name}:\s*([^\n]+)"
                        match = re.search(pattern, docstring)
                        if match:
                            param_schema["description"] = match.group(1).strip()
                        else:
                            param_schema["description"] = f"Parameter: {param_name}"
                    else:
                        param_schema["description"] = f"Parameter: {param_name}"
                    
                    properties[param_name] = param_schema
                
                schema = {
                    "type": "object",
                    "properties": properties
                }
                
                if required:
                    schema["required"] = required
                
                print(f"  Extracted {len(properties)} properties for {func_name}")
                return schema
            
            # Function not found in AST - list all function names for debugging
            all_funcs = []
            for node in ast.walk(tree):
                if isinstance(node, ast.FunctionDef):
                    all_funcs.append(node.name)
            print(f"Warning: Function {func_name} not found in AST of {source_file}")
            print(f"  Available functions in file: {all_funcs[:10]}...")  # First 10
            return {"type": "object", "properties": {}}
            
        except Exception as e:
            print(f"Error extracting schema from source for {func_name} in {inspect.getfile(module) if hasattr(module, '__file__') else 'unknown'}: {e}")
            import traceback
            traceback.print_exc()
            return {"type": "object", "properties": {}}
    
    def _get_tool_schema(self, tool_func):
        """Extract JSON schema from tool function signature."""
        if not callable(tool_func):
            print(f"  ERROR: tool_func is not callable: {type(tool_func)}", flush=True)
            return {"type": "object", "properties": {}}
        
        # Unwrap if function is wrapped (e.g., by FastMCP decorator)
        original_func = tool_func
        unwrap_attempts = 0
        max_unwrap = 10  # Increased to handle more wrapping layers
        
        while unwrap_attempts < max_unwrap:
            if hasattr(tool_func, '__wrapped__'):
                tool_func = tool_func.__wrapped__
                unwrap_attempts += 1
            elif hasattr(tool_func, '_func'):
                tool_func = tool_func._func
                unwrap_attempts += 1
            elif hasattr(tool_func, 'func'):
                tool_func = tool_func.func
                unwrap_attempts += 1
            elif hasattr(tool_func, '__func__'):
                tool_func = tool_func.__func__
                unwrap_attempts += 1
            else:
                break
        
        # If we couldn't unwrap, try the original
        if not callable(tool_func):
            tool_func = original_func
        
        # Try to get signature
        sig = None
        try:
            sig = inspect.signature(tool_func)
            print(f"  DEBUG: Got signature for function: {len(sig.parameters)} parameters", flush=True)
        except (ValueError, TypeError) as e:
            print(f"  DEBUG: inspect.signature() failed: {e}", flush=True)
            # If signature extraction fails, try to get it from source code
            try:
                source = inspect.getsource(tool_func)
                # Try to parse AST to get function signature
                tree = ast.parse(source)
                for node in ast.walk(tree):
                    if isinstance(node, ast.FunctionDef):
                        # Found the function definition
                        # Create a mock signature from AST
                        params = []
                        for arg in node.args.args:
                            if arg.arg != 'self':
                                param_name = arg.arg
                                # Try to infer type from annotation
                                param_type = str
                                if arg.annotation:
                                    if isinstance(arg.annotation, ast.Name):
                                        type_name = arg.annotation.id
                                        if type_name == 'int':
                                            param_type = int
                                        elif type_name == 'float':
                                            param_type = float
                                        elif type_name == 'bool':
                                            param_type = bool
                                        elif type_name == 'list' or type_name == 'List':
                                            param_type = list
                                        elif type_name == 'dict' or type_name == 'Dict':
                                            param_type = dict
                                    elif isinstance(arg.annotation, ast.Constant):
                                        param_type = type(arg.annotation.value)
                                
                                # Check if it's Optional
                                is_optional = False
                                if isinstance(arg.annotation, ast.Subscript):
                                    if isinstance(arg.annotation.value, ast.Name) and arg.annotation.value.id == 'Optional':
                                        is_optional = True
                                
                                default_value = inspect.Parameter.empty
                                # Check for default value in AST (would need to match args with defaults)
                                params.append((param_name, param_type, is_optional, default_value))
                        
                        # Create a mock signature object
                        class MockParam:
                            def __init__(self, name, annotation, default):
                                self.name = name
                                self.annotation = annotation
                                self.default = default
                        
                        class MockSig:
                            def __init__(self, params_list):
                                self.parameters = {}
                                for i, (name, ann, opt, default) in enumerate(params_list):
                                    # Adjust for defaults
                                    if i >= len(params_list) - len(node.args.defaults):
                                        default_idx = i - (len(params_list) - len(node.args.defaults))
                                        if default_idx < len(node.args.defaults):
                                            default = inspect.Parameter.empty  # Would need to evaluate default
                                    self.parameters[name] = MockParam(name, ann, default)
                        
                        sig = MockSig(params)
                        break
            except Exception as ast_error:
                print(f"Warning: Could not get signature from source either: {ast_error}")
                return {"type": "object", "properties": {}}
        
        if sig is None:
            return {"type": "object", "properties": {}}
        
        properties = {}
        required = []
        
        # Get docstring for better descriptions
        docstring = tool_func.__doc__ or ""
        
        for param_name, param in sig.parameters.items():
            if param_name == 'self':
                continue
            
            param_type = param.annotation
            param_default = param.default
            
            # Handle Optional types and Union types
            if param_type != inspect.Parameter.empty:
                # Check if it's a Union type (Optional is Union[T, None])
                if hasattr(typing, 'get_origin'):
                    origin = typing.get_origin(param_type)
                    if origin is typing.Union:
                        args = typing.get_args(param_type)
                        # If Union includes None, it's Optional - get the actual type
                        non_none_args = [arg for arg in args if arg is not type(None)]
                        if non_none_args:
                            param_type = non_none_args[0]
                    elif origin is list:
                        # Handle List[T] types
                        args = typing.get_args(param_type)
                        if args:
                            param_type = list  # We'll handle the item type separately
                # Also check for string representation (for older Python versions)
                elif isinstance(param_type, str):
                    if 'Optional' in str(param_type) or 'Union' in str(param_type):
                        # Try to extract the base type
                        if 'int' in str(param_type):
                            param_type = int
                        elif 'float' in str(param_type):
                            param_type = float
                        elif 'bool' in str(param_type):
                            param_type = bool
                        elif 'list' in str(param_type) or 'List' in str(param_type):
                            param_type = list
                        elif 'dict' in str(param_type) or 'Dict' in str(param_type):
                            param_type = dict
                        else:
                            param_type = str
            
            # Map Python types to JSON schema types
            prop_schema = {}
            
            if param_type == inspect.Parameter.empty or param_type == type(None):
                prop_schema["type"] = "string"
            elif param_type == str:
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
                # Default to string for unknown types
                prop_schema["type"] = "string"
                print(f"    DEBUG: Unknown param type {param_type} for {param_name}, defaulting to string", flush=True)
            
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

