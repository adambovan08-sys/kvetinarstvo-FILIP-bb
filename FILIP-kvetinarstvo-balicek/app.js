const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwOy8TvOnBk0Cq01ZIcV-3fDtsRgjJhMcVKl5cUH4KaclrKV3Lxo0tv6XfQKi_hklhL/exec";

const header = document.querySelector(".site-header");
const hero = document.querySelector(".hero");
const form = document.querySelector("#flowerOrderForm");
const formStatus = document.querySelector("#formStatus");
const submitButton = document.querySelector(".submit-button");
const year = document.querySelector("#currentYear");

form.action = GOOGLE_SCRIPT_URL;
form.method = "GET";
form.acceptCharset = "UTF-8";

year.textContent = new Date().getFullYear();

const setHeaderState = () => {
  header.classList.toggle("is-scrolled", window.scrollY > 18);
};

window.addEventListener("scroll", setHeaderState, { passive: true });
setHeaderState();

document.querySelectorAll("[data-choice]").forEach((button) => {
  button.addEventListener("click", () => {
    const choice = button.dataset.choice;
    form.elements.bouquetStyle.value = choice;
    document.querySelector("#objednavka").scrollIntoView({ behavior: "smooth", block: "start" });
    form.elements.budget.focus({ preventScroll: true });
  });
});

const prefersReducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

if (!prefersReducedMotion) {
  hero.addEventListener(
    "pointermove",
    (event) => {
      const rect = hero.getBoundingClientRect();
      const x = (event.clientX - rect.left) / rect.width - 0.5;
      const y = (event.clientY - rect.top) / rect.height - 0.5;
      hero.style.setProperty("--hero-x", `${x * -18}px`);
      hero.style.setProperty("--hero-y", `${y * -14}px`);
    },
    { passive: true }
  );

  document.querySelectorAll(".motion-card").forEach((card) => {
    const reset = () => {
      card.style.setProperty("--tilt-x", "0deg");
      card.style.setProperty("--tilt-y", "0deg");
      card.style.setProperty("--shift-x", "0px");
      card.style.setProperty("--shift-y", "0px");
      card.classList.remove("is-pressed");
    };

    card.addEventListener(
      "pointermove",
      (event) => {
        const rect = card.getBoundingClientRect();
        const x = (event.clientX - rect.left) / rect.width - 0.5;
        const y = (event.clientY - rect.top) / rect.height - 0.5;
        card.style.setProperty("--tilt-x", `${y * -8}deg`);
        card.style.setProperty("--tilt-y", `${x * 8}deg`);
        card.style.setProperty("--shift-x", `${x * 4}px`);
        card.style.setProperty("--shift-y", `${y * 4}px`);
      },
      { passive: true }
    );

    card.addEventListener("pointerdown", () => card.classList.add("is-pressed"), { passive: true });
    card.addEventListener("pointerup", reset, { passive: true });
    card.addEventListener("pointercancel", reset, { passive: true });
    card.addEventListener("pointerleave", reset, { passive: true });
  });
}

const revealTargets = document.querySelectorAll(
  ".hero-copy > *, .quick-panel, .service-strip article, .story-layout > *, .section-heading, .bouquet-card, .product-card, .order-copy, .order-form, .contact-layout > *"
);

revealTargets.forEach((element, index) => {
  element.classList.add("reveal");
  element.style.setProperty("--reveal-delay", `${Math.min(index % 6, 5) * 70}ms`);
});

if ("IntersectionObserver" in window && !prefersReducedMotion) {
  const revealObserver = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          entry.target.classList.add("is-visible");
          revealObserver.unobserve(entry.target);
        }
      });
    },
    {
      threshold: 0.16,
      rootMargin: "0px 0px -8% 0px",
    }
  );

  revealTargets.forEach((element) => revealObserver.observe(element));
} else {
  revealTargets.forEach((element) => element.classList.add("is-visible"));
}

const setStatus = (message, isError = false) => {
  formStatus.textContent = message;
  formStatus.classList.toggle("is-error", isError);
};

const setHiddenField = (name, value) => {
  let input = form.querySelector(`input[type="hidden"][name="${name}"]`);
  if (!input) {
    input = document.createElement("input");
    input.type = "hidden";
    input.name = name;
    form.appendChild(input);
  }
  input.value = value;
};

form.addEventListener("submit", (event) => {
  if (!form.reportValidity()) {
    event.preventDefault();
    return;
  }

  if (!GOOGLE_SCRIPT_URL) {
    event.preventDefault();
    setStatus("Formulár je pripravený. Doplňte Google Apps Script URL v app.js.", true);
    return;
  }

  setHiddenField("createdAt", new Date().toISOString());
  setHiddenField("source", window.location.href);

  submitButton.disabled = true;
  setStatus("Odosielam objednávku...");
});
