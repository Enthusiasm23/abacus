from abacus.mcp_server import mcp


class TestMcpServer:
    def test_server_name(self):
        assert mcp.name == "abacus"