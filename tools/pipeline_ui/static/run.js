/**
 * run.js -- SSE EventSource client for the live pipeline log page.
 *
 * Connects to /stream/<job_id>, receives log lines one at a time,
 * appends them to the log body, advances the step progress pips,
 * and handles checkpoint and done events.
 */

function startStream(jobId) {
  const logBody    = document.getElementById("logBody");
  const logBadge   = document.getElementById("logBadge");
  const statusLabel = document.getElementById("statusLabel");
  const stepBar    = document.getElementById("stepBar");

  let currentStep  = 0;
  let checkpointPending = false;

  const source = new EventSource(`/stream/${jobId}`);

  source.onmessage = function(e) {
    const msg = JSON.parse(e.data);

    // ── Heartbeat — keep connection alive ──────────────────────────────
    if (msg.type === "heartbeat") return;

    // ── Log line ────────────────────────────────────────────────────────
    if (msg.type === "log") {
      // Advance step pip when step number changes.
      if (msg.step && msg.step !== currentStep) {
        currentStep = msg.step;
        markStep(currentStep, "active");
        if (currentStep > 1) markStep(currentStep - 1, "done");
        statusLabel.textContent = `Running step ${currentStep} of 10…`;
      }
      appendLine(msg.line, "log-line");
    }

    // ── Checkpoint ──────────────────────────────────────────────────────
    if (msg.type === "checkpoint") {
      source.close();
      checkpointPending = true;
      logBadge.textContent = "Checkpoint";
      logBadge.className   = "log-badge checkpoint";
      statusLabel.textContent = "Review required — pipeline paused";
      appendLine("─── CHECKPOINT: Steps 1–2 complete. Review the Colored Connections Table. ───", "log-separator");

      // Redirect to checkpoint page after a short delay.
      setTimeout(() => {
        window.location.href = `/checkpoint/${jobId}`;
      }, 1200);
    }

    // ── Done ────────────────────────────────────────────────────────────
    if (msg.type === "done") {
      source.close();

      if (msg.status === "complete") {
        markStep(currentStep, "done");
        logBadge.textContent = "Complete";
        logBadge.className   = "log-badge done";
        statusLabel.textContent = "All steps finished — preparing downloads…";
        appendLine("─── Pipeline complete ───", "log-separator");
        setTimeout(() => {
          window.location.href = `/complete/${jobId}`;
        }, 1500);
      } else if (msg.status === "failed") {
        markStep(currentStep, "failed");
        logBadge.textContent = "Failed";
        logBadge.className   = "log-badge failed";
        statusLabel.textContent = "Pipeline failed — see log for details";
        appendLine(`ERROR: ${msg.error || "Unknown error"}`, "log-error");
      } else {
        // stopped_at_checkpoint or other
        window.location.href = `/complete/${jobId}`;
      }
    }
  };

  source.onerror = function() {
    if (!checkpointPending) {
      logBadge.textContent = "Disconnected";
      logBadge.className   = "log-badge failed";
      appendLine("Connection to server lost. Check that the pipeline is still running.", "log-error");
    }
  };

  // ── Helpers ─────────────────────────────────────────────────────────────

  function appendLine(text, cls) {
    if (!text && !text.trim()) return;
    const line = document.createElement("div");
    line.className = cls || "log-line";
    line.textContent = text;
    logBody.appendChild(line);
    logBody.scrollTop = logBody.scrollHeight;
  }

  function markStep(stepNum, state) {
    const pip = document.getElementById(`pip-${stepNum}`);
    if (!pip) return;
    pip.className = `step-pip ${state}`;
  }
}
