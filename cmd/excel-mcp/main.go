package main

import (
	"context"
	"fmt"
	"log/slog"
	"net/http"
	"os"
	"time"

	"excel-mcp/internal/cli"
	"excel-mcp/internal/server"

	"github.com/modelcontextprotocol/go-sdk/mcp"
)

func main() {
	logger := slog.New(slog.NewTextHandler(os.Stderr, &slog.HandlerOptions{Level: slog.LevelInfo}))
	ctx := context.Background()

	mode := "stdio"
	if len(os.Args) > 1 {
		mode = os.Args[1]
	}

	switch mode {
	case "cli":
		os.Exit(cli.Run(ctx, os.Args[2:], os.Stdout, os.Stderr, logger))
	case "stdio":
		srv := server.New(server.Config{PathMode: server.PathModeDirect, Logger: logger})
		if err := srv.Run(ctx, &mcp.StdioTransport{}); err != nil {
			logger.Error("server exited", "error", err)
			os.Exit(1)
		}
	case "streamable-http":
		port := os.Getenv("EXCEL_MCP_SERVER_PORT")
		if port == "" {
			port = "8000"
		}
		handler := mcp.NewStreamableHTTPHandler(func(*http.Request) *mcp.Server {
			return server.New(server.Config{PathMode: server.PathModeRooted, Logger: logger})
		}, nil)
		addr := ":" + port
		logger.Info("starting streamable HTTP server", "addr", addr, "path", "/mcp")
		mux := http.NewServeMux()
		mux.Handle("/mcp", handler)
		srv := &http.Server{
			Addr:         addr,
			Handler:      mux,
			ReadTimeout:  30 * time.Second,
			WriteTimeout: 30 * time.Second,
		}
		if err := srv.ListenAndServe(); err != nil {
			logger.Error("server exited", "error", err)
			os.Exit(1)
		}
	default:
		fmt.Fprintf(os.Stderr, "usage: %s [cli|stdio|streamable-http]\n", os.Args[0])
		os.Exit(2)
	}
}
