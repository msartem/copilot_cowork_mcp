"""Entry point for `python -m copilot_cowork_mcp` and console script."""

from copilot_cowork_mcp.server import mcp


def main():
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
