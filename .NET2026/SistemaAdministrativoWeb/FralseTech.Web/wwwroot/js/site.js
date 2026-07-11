const header = document.querySelector("[data-header]");
const menuToggle = document.querySelector("[data-menu-toggle]");
const navigation = document.querySelector("[data-navigation]");
const navLinks = navigation ? navigation.querySelectorAll("a") : [];
const prefersReducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

if (header) {
    const updateHeaderState = () => {
        header.classList.toggle("is-scrolled", window.scrollY > 12);
    };

    updateHeaderState();
    window.addEventListener("scroll", updateHeaderState, { passive: true });
}

if (menuToggle && navigation) {
    const closeMenu = () => {
        navigation.classList.remove("is-open");
        menuToggle.setAttribute("aria-expanded", "false");
    };

    menuToggle.addEventListener("click", () => {
        const isOpen = navigation.classList.toggle("is-open");
        menuToggle.setAttribute("aria-expanded", isOpen ? "true" : "false");
    });

    navLinks.forEach((link) => {
        link.addEventListener("click", () => {
            if (window.innerWidth <= 840) {
                closeMenu();
            }
        });
    });

    window.addEventListener("resize", () => {
        if (window.innerWidth > 840) {
            closeMenu();
        }
    });
}

if (!prefersReducedMotion) {
    const revealElements = document.querySelectorAll(".reveal");
    const observer = new IntersectionObserver((entries) => {
        entries.forEach((entry) => {
            if (entry.isIntersecting) {
                entry.target.classList.add("is-visible");
                observer.unobserve(entry.target);
            }
        });
    }, { threshold: 0.18 });

    revealElements.forEach((element) => observer.observe(element));
} else {
    document.querySelectorAll(".reveal").forEach((element) => element.classList.add("is-visible"));
}
