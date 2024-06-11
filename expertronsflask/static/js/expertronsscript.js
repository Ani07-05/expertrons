document.addEventListener('DOMContentLoaded', function () {
    const form = document.getElementById('registration-form');

    form.addEventListener('submit', function (event) {
        event.preventDefault(); // Prevent form submission

        // Validate form fields
        if (validateForm()) {
            // Display success message
            showAlert('Success! Registration submitted successfully.');
            // Optionally, submit form data to server
            // form.submit();
        }
    });

    function validateForm() {
        // Perform form validation here
        // For simplicity, let's assume all fields are required
        const inputs = form.querySelectorAll('input, select');
        let isValid = true;

        inputs.forEach(input => {
            if (!input.value) {
                isValid = false;
                input.classList.add('error');
            } else {
                input.classList.remove('error');
            }
        });

        return isValid;
    }

    function showAlert(message) {
        const alert = document.createElement('div');
        alert.className = 'alert success';
        alert.textContent = message;
        document.body.appendChild(alert);
        setTimeout(function () {
            alert.remove();
        }, 3000); // Remove alert after 3 seconds
    }
});
document.addEventListener('DOMContentLoaded', function() {
    const counters = document.querySelectorAll('.floating-text');

    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                const counter = entry.target;
                const finalNumber = parseInt(counter.getAttribute('data-target').replace(/\D/g, ''));
                let currentNumber = 0;

                const increment = Math.ceil(finalNumber / 100); // Adjust the divisor to control speed

                const updateCounter = () => {
                    currentNumber += increment;
                    if (currentNumber < finalNumber) {
                        counter.innerText = `${numberWithCommas(currentNumber)}+`;
                        setTimeout(updateCounter, 50); // Adjust timing for faster/slower animation
                    } else {
                        counter.innerText = `${numberWithCommas(finalNumber)}+`;
                    }
                };
                
                updateCounter();
                observer.unobserve(counter); // Stop observing once animated
            }
        });
    }, {
        threshold: 0.5 // Trigger when 50% of the element is visible
    });

    counters.forEach(counter => {
        observer.observe(counter);
    });

    function numberWithCommas(x) {
        return x.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    }
});

