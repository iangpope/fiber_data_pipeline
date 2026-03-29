/**
 * run.js -- SSE EventSource client for the live pipeline log page.
 *
 * Normal operation: shows a spinner + step label only.
 * On failure: reveals the log card with buffered output for debugging.
 */

function startStream(jobId) {
  const processingArea  = document.getElementById("processingArea");
  const processingLabel = document.getElementById("processingLabel");
  const logCard         = document.getElementById("logCard");
  const logBody         = document.getElementById("logBody");
  const statusLabel     = document.getElementById("statusLabel");

  let currentStep       = 0;
  let checkpointPending = false;
  const bufferedLines   = [];   // all log lines, kept in case we need to show them

  const STEP_NAMES = {
    1: "Extracting KMZ coordinates",
    2: "Computing cable directions",
    3: "Colorizing cut sheet",
    4: "Formatting top section",
    5: "Normalizing connections",
    6: "Assigning addresses",
    7: "Processing taps",
    8: "Finalizing workbook",
    9: "Tracing path of light",
    10: "Generating tap report",
  };

  const source = new EventSource(`/stream/${jobId}`);

  source.onmessage = function(e) {
    const msg = JSON.parse(e.data);

    // ── Heartbeat — keep connection alive ──────────────────────────────
    if (msg.type === "heartbeat") return;

    // ── Log line ────────────────────────────────────────────────────────
    if (msg.type === "log") {
      bufferedLines.push(msg.line);

      if (msg.step && msg.step !== currentStep) {
        currentStep = msg.step;
        markStep(currentStep, "active");
        if (currentStep > 1) markStep(currentStep - 1, "done");
        const name = STEP_NAMES[currentStep] || `Step ${currentStep}`;
        statusLabel.textContent  = `Step ${currentStep} of 10`;
        processingLabel.textContent = name + "…";
      }
    }

    // ── Checkpoint ──────────────────────────────────────────────────────
    if (msg.type === "checkpoint") {
      source.close();
      checkpointPending = true;
      statusLabel.textContent      = "Checkpoint — pipeline paused";
      processingLabel.textContent  = "Review required…";
      setTimeout(() => {
        window.location.href = `/checkpoint/${jobId}`;
      }, 1200);
    }

    // ── Done ────────────────────────────────────────────────────────────
    if (msg.type === "done") {
      source.close();

      if (msg.status === "complete") {
        markStep(currentStep, "done");
        statusLabel.textContent     = "All steps finished";
        processingLabel.textContent = "Complete — preparing downloads…";
        setTimeout(() => {
          window.location.href = `/complete/${jobId}`;
        }, 1500);

      } else if (msg.status === "failed") {
        markStep(currentStep, "failed");
        statusLabel.textContent = "Pipeline failed";

        // Reveal the log card and dump everything we buffered
        processingArea.style.display = "none";
        logCard.style.display        = "block";
        bufferedLines.forEach(l => appendLine(l, "log-line"));
        appendLine(`ERROR: ${msg.error || "Unknown error"}`, "log-error");

      } else {
        // stopped_at_checkpoint or other
        window.location.href = `/complete/${jobId}`;
      }
    }
  };

  source.onerror = function() {
    if (!checkpointPending) {
      processingArea.style.display = "none";
      logCard.style.display        = "block";
      bufferedLines.forEach(l => appendLine(l, "log-line"));
      appendLine("Connection to server lost. Check that the pipeline is still running.", "log-error");
    }
  };

  // ── Helpers ─────────────────────────────────────────────────────────────

  function appendLine(text, cls) {
    if (!text || !text.trim()) return;
    const line = document.createElement("div");
    line.className  = cls || "log-line";
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
