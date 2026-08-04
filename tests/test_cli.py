from click.testing import CliRunner
from abacus.cli import main


class TestCli:
    def test_main_command(self):
        runner = CliRunner()
        result = runner.invoke(main, ["--help"])
        assert result.exit_code == 0
        assert "Abacus" in result.output
    
    def test_capabilities_command(self):
        runner = CliRunner()
        result = runner.invoke(main, ["capabilities"])
        assert result.exit_code == 0