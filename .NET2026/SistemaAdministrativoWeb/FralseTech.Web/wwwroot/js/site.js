const header = document.querySelector("[data-header]");
const menuToggle = document.querySelector("[data-menu-toggle]");
const navigation = document.querySelector("[data-navigation]");
const navLinks = navigation ? navigation.querySelectorAll("a") : [];
const prefersReducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

const updateHeaderState = () => {
    header?.classList.toggle("is-scrolled", window.scrollY > 18);
};

const setMenuState = (isOpen) => {
    if (!menuToggle || !navigation) return;

    navigation.classList.toggle("is-open", isOpen);
    menuToggle.setAttribute("aria-expanded", String(isOpen));
    document.body.style.overflow = isOpen ? "hidden" : "";
};

updateHeaderState();
window.addEventListener("scroll", updateHeaderState, { passive: true });

if (menuToggle && navigation) {
    menuToggle.addEventListener("click", () => {
        setMenuState(!navigation.classList.contains("is-open"));
    });

    navLinks.forEach((link) => link.addEventListener("click", () => setMenuState(false)));

    window.addEventListener("keydown", (event) => {
        if (event.key === "Escape") setMenuState(false);
    });

    window.addEventListener("resize", () => {
        if (window.innerWidth > 820) setMenuState(false);
    });
}

if (!prefersReducedMotion) {
    const observer = new IntersectionObserver((entries) => {
        entries.forEach((entry) => {
            if (!entry.isIntersecting) return;
            entry.target.classList.add("is-visible");
            observer.unobserve(entry.target);
        });
    }, { threshold: 0.14 });

    document.querySelectorAll(".reveal").forEach((element) => observer.observe(element));
} else {
    document.querySelectorAll(".reveal").forEach((element) => element.classList.add("is-visible"));
}
