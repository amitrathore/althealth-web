const formConfigs = {
  form01: {
    recipient: "webforms@althealth.store",
    subject: "AltHealth Playbook Subscription",
  },
  form02: {
    recipient: "webforms@althealth.store",
    subject: "AltHealth Consultation Request",
  },
  form03: {
    recipient: "webforms@althealth.store",
    subject: "AltHealth Brand Application",
  },
  form04: {
    recipient: "webforms@althealth.store",
    subject: "AltHealth Venture Network Application",
  },
};

function setSubmitting(form, submitting) {
  const button = form.querySelector('button[type="submit"]');
  if (!button) return;
  button.disabled = submitting;
  button.dataset.label ||= button.textContent;
  button.textContent = submitting ? "Sending..." : button.dataset.label;
}

function buildMailto(form, config) {
  const data = new FormData(form);
  const lines = [];

  for (const [key, value] of data.entries()) {
    if (key === "id") continue;
    lines.push(`${key}: ${value}`);
  }

  return `mailto:${config.recipient}?subject=${encodeURIComponent(config.subject)}&body=${encodeURIComponent(lines.join("\n"))}`;
}

function submitForm(form, config) {
  if (!form.reportValidity()) return;

  setSubmitting(form, true);

  try {
    const mailto = buildMailto(form, config);
    window.location.href = mailto;
    form.reset();

    const successEl = document.getElementById(form.id + "-success");
    if (successEl) {
      form.style.display = "none";
      successEl.classList.add("visible");
    }
  } catch (error) {
    window.alert(error.message || "Sorry, something went wrong. Please try again later.");
  } finally {
    setSubmitting(form, false);
  }
}

Object.entries(formConfigs).forEach(([id, config]) => {
  const form = document.getElementById(id);
  if (!form) return;

  form.addEventListener("submit", (event) => {
    event.preventDefault();
    submitForm(form, config);
  });
});
