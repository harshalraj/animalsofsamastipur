/* ==========================================================================
   Animals of Samastipur - Client JavaScript Logic
   ========================================================================== */

document.addEventListener('DOMContentLoaded', () => {

    // --- Mobile Menu Toggle ---
    const menuToggle = document.getElementById('menuToggle');
    const navLinks = document.getElementById('navLinks');
    if (menuToggle && navLinks) {
        menuToggle.addEventListener('click', () => {
            navLinks.classList.toggle('active');
        });
    }

    // --- Animal Care Category Filter Tabs ---
    const filterBtns = document.querySelectorAll('.filter-btn');
    const updateCards = document.querySelectorAll('.update-card');

    filterBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            filterBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');

            const category = btn.dataset.category;
            updateCards.forEach(card => {
                if (category === 'all' || card.dataset.category === category) {
                    card.style.display = 'flex';
                } else {
                    card.style.display = 'none';
                }
            });
        });
    });

    // --- Astrology Form Direct API Submission ---
    const astrologyForm = document.getElementById('astrologyForm');
    const submitAstroBtn = document.getElementById('submitAstroBtn');

    if (astrologyForm) {
        astrologyForm.addEventListener('submit', async (e) => {
            e.preventDefault();

            const formData = {
                name: document.getElementById('userName').value.trim(),
                email: document.getElementById('userEmail').value.trim(),
                dob: document.getElementById('userDob').value.trim(),
                tob: document.getElementById('userTob').value.trim(),
                pob: document.getElementById('userPob').value.trim(),
                question: document.getElementById('userQuestion').value.trim()
            };

            if (!formData.name || !formData.email || !formData.question) {
                showToast('Please fill all required fields.', 'error');
                return;
            }

            const originalText = submitAstroBtn.innerHTML;
            submitAstroBtn.disabled = true;
            submitAstroBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Submitting Query...';

            try {
                const response = await fetch('/api/submit-astrology', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(formData)
                });

                const result = await response.json();

                if (result.success) {
                    showToast(result.message, 'success');
                    astrologyForm.reset();
                } else {
                    showToast(result.message || 'Submission failed.', 'error');
                }
            } catch (err) {
                showToast('Server connection error. Please try again.', 'error');
            } finally {
                submitAstroBtn.disabled = false;
                submitAstroBtn.innerHTML = originalText;
            }
        });
    }

    // --- Volunteer Form Submission ---
    const volunteerForm = document.getElementById('volunteerForm');
    if (volunteerForm) {
        volunteerForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const submitBtn = volunteerForm.querySelector('button[type="submit"]');
            const originalText = submitBtn.innerHTML;

            submitBtn.disabled = true;
            submitBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Submitting...';

            const formData = new FormData(volunteerForm);
            const dataObj = Object.fromEntries(formData.entries());

            try {
                const response = await fetch('/api/volunteer', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(dataObj)
                });

                const result = await response.json();

                if (result.success) {
                    showToast(result.message, 'success');
                    volunteerForm.reset();
                } else {
                    showToast(result.message || 'Volunteer registration failed.', 'error');
                }
            } catch (err) {
                showToast('Error connecting to server.', 'error');
            } finally {
                submitBtn.disabled = false;
                submitBtn.innerHTML = originalText;
            }
        });
    }

    // --- Toast Notification Helper ---
    function showToast(message, type = 'success') {
        const toastContainer = document.getElementById('toastContainer');
        if (!toastContainer) return;

        const toast = document.createElement('div');
        toast.className = `toast ${type === 'error' ? 'toast-error' : ''}`;
        
        const icon = type === 'error' ? 'fa-triangle-exclamation' : 'fa-circle-check';
        toast.innerHTML = `<i class="fa-solid ${icon}"></i> <span>${message}</span>`;

        toastContainer.appendChild(toast);

        setTimeout(() => {
            toast.style.opacity = '0';
            toast.style.transform = 'translateX(100%)';
            toast.style.transition = 'all 0.4s ease';
            setTimeout(() => toast.remove(), 400);
        }, 5000);
    }
});
