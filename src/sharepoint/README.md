# Redis

A Model Context Protocol server that provides access to Organisational Sharepoint.

## Components

### Tools
- Connects to Sharepoint using Microsoft Graph API
- Exposes Sharepoint documents as resources
- Provides tools for searching documents
- Includes prompts for common Sharepoint tasks


## Usage with Claude Desktop

To use this server with the Claude Desktop app, add the following configuration to the "mcpServers" section of your `claude_desktop_config.json`:

### Docker

* when running docker on macos, use host.docker.internal if the server is running on the host network (eg localhost)
* Sharepoint URL can be specified as an argument, defaults to "redis://localhost:6379"

```json
{
  "mcpServers": {
    "redis": {
      "command": "docker",
      "args": [
        "run", 
        "-i", 
        "--rm", 
        "mcp/sharepoint"
        ]
    }
  }
}
```

### NPX

```json
{
  "mcpServers": {
    "redis": {
      "command": "npx",
      "args": [
        "-y",
        "@modelcontextprotocol/server-sharepoint"
      ]
    }
  }
}
```

## Building

Docker:

```sh
docker build -t mcp/sharepoint -f src/sharepoint/Dockerfile . 
```

## License

This MCP server is licensed under the MIT License. This means you are free to use, modify, and distribute the software, subject to the terms and conditions of the MIT License. For more details, please see the LICENSE file in the project repository.