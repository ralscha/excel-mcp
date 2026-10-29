package cli

import (
	"bytes"
	"testing"
)

func TestSubcommandHelpSucceeds(t *testing.T) {
	for _, command := range []string{"list-tools", "tool-info", "create_workbook"} {
		t.Run(command, func(t *testing.T) {
			var stdout, stderr bytes.Buffer
			if code := Run(t.Context(), []string{command, "--help"}, &stdout, &stderr, testLogger()); code != 0 {
				t.Fatalf("help exited with %d: %s", code, stderr.String())
			}
			if stdout.Len()+stderr.Len() == 0 {
				t.Fatal("expected help text")
			}
		})
	}
}

func TestInvalidSubcommandFlagStillFails(t *testing.T) {
	for _, command := range []string{"list-tools", "tool-info", "create_workbook"} {
		t.Run(command, func(t *testing.T) {
			var stdout, stderr bytes.Buffer
			if code := Run(t.Context(), []string{command, "--not-a-flag"}, &stdout, &stderr, testLogger()); code != 2 {
				t.Fatalf("invalid flag exited with %d: %s", code, stderr.String())
			}
		})
	}
}
