document.addEventListener("DOMContentLoaded", function() {
    // Select all input fields within the form that should be navigable
    const form = document.querySelector(".validated-form");
    if (!form) return; // Exit if there's no form on the page

    const inputs = Array.from(form.querySelectorAll('input[type="number"]'));

    if (inputs.length === 0) return; // Exit if no navigable inputs are found

    inputs.forEach((input, index) => {
        input.addEventListener("keydown", (e) => {
            // Handle 'Enter' and 'ArrowDown' to move to the next input
            if (e.key === "Enter" || e.key === "ArrowDown") {
                e.preventDefault(); // Prevent form submission on Enter
                const nextInput = inputs[index + 1];
                if (nextInput) {
                    nextInput.focus();
                    nextInput.select(); // Highlight the content of the next input
                }
            }
            // Handle 'ArrowUp' to move to the previous input
            else if (e.key === "ArrowUp") {
                e.preventDefault();
                const prevInput = inputs[index - 1];
                if (prevInput) {
                    prevInput.focus();
                    prevInput.select();
                }
            }
        });
    });
});