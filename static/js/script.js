document.addEventListener('DOMContentLoaded', function() {
    // Get form elements
    const uploadForm = document.getElementById('upload-form');
    const fileInput = document.getElementById('file');
    const submitBtn = document.getElementById('submit-btn');
    const progressContainer = document.getElementById('progress-container');
    const processingMessage = document.getElementById('processing-message');
    
    // Add form submit handler
    if (uploadForm) {
        uploadForm.addEventListener('submit', function(e) {
            // Check if a file is selected
            if (fileInput.files.length === 0) {
                e.preventDefault();
                showAlert('Please select an Excel file (.xlsx) to upload.', 'danger');
                return;
            }
            
            // Validate file type
            const file = fileInput.files[0];
            const fileExtension = file.name.split('.').pop().toLowerCase();
            
            if (fileExtension !== 'xlsx') {
                e.preventDefault();
                showAlert('Only .xlsx files are accepted.', 'danger');
                return;
            }
            
            // Show processing UI
            submitBtn.disabled = true;
            progressContainer.classList.remove('d-none');
            processingMessage.classList.remove('d-none');
        });
    }

    // File input change handler to validate file type
    if (fileInput) {
        fileInput.addEventListener('change', function() {
            if (this.files.length > 0) {
                const file = this.files[0];
                const fileExtension = file.name.split('.').pop().toLowerCase();
                
                if (fileExtension !== 'xlsx') {
                    showAlert('Only .xlsx files are accepted.', 'warning');
                    this.value = ''; // Clear the file input
                }
            }
        });
    }

    // Function to show alerts
    function showAlert(message, type) {
        // Create alert element
        const alertDiv = document.createElement('div');
        alertDiv.className = `alert alert-${type} alert-dismissible fade show`;
        alertDiv.role = 'alert';
        
        // Add message and close button
        alertDiv.innerHTML = `
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
        `;
        
        // Find the location to insert the alert (before the form)
        const form = document.getElementById('upload-form');
        form.parentNode.insertBefore(alertDiv, form);
        
        // Auto-close after 5 seconds
        setTimeout(() => {
            alertDiv.classList.remove('show');
            setTimeout(() => {
                alertDiv.remove();
            }, 150);
        }, 5000);
    }

    // Initialize Bootstrap tooltips
    const tooltipTriggerList = document.querySelectorAll('[data-bs-toggle="tooltip"]');
    const tooltipList = [...tooltipTriggerList].map(tooltipTriggerEl => new bootstrap.Tooltip(tooltipTriggerEl));
});
