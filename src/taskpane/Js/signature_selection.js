// Signature selection interactive functionality

document.addEventListener('DOMContentLoaded', function() {
    // Initialize the signature selection functionality
    initializeSignatureSelection();
});

function initializeSignatureSelection() {
    const checkboxes = document.querySelectorAll('.option-card input[type="checkbox"]');
    const selectedOptionsContainer = document.getElementById('selectedOptions');
    
    // Add event listeners to all checkboxes
    checkboxes.forEach(checkbox => {
        // Set initial state
        updateCardState(checkbox);
        updateSelectedDisplay();
        
        // Add change event listener
        checkbox.addEventListener('change', function() {
            updateCardState(this);
            updateSelectedDisplay();
        });
    });
    
    // Add click handlers to cards for better UX
    const cards = document.querySelectorAll('.option-card');
    cards.forEach(card => {
        card.addEventListener('click', function(e) {
            // Prevent double-triggering when clicking directly on checkbox
            if (e.target.type === 'checkbox') return;
            
            const checkbox = this.querySelector('input[type="checkbox"]');
            checkbox.checked = !checkbox.checked;
            checkbox.dispatchEvent(new Event('change'));
        });
    });
}

function updateCardState(checkbox) {
    const card = checkbox.closest('.option-card');
    
    if (checkbox.checked) {
        card.classList.add('selected');
    } else {
        card.classList.remove('selected');
    }
}

function updateSelectedDisplay() {
    const selectedOptionsContainer = document.getElementById('selectedOptions');
    const checkboxes = document.querySelectorAll('.option-card input[type="checkbox"]:checked');
    
    // Clear existing selected items
    selectedOptionsContainer.innerHTML = '';
    
    // Add selected items with animation
    checkboxes.forEach((checkbox, index) => {
        const label = checkbox.closest('.option-card').querySelector('.option-title').textContent;
        const selectedItem = createSelectedItem(label);
        selectedOptionsContainer.appendChild(selectedItem);
        
        // Trigger animation after a small delay
        setTimeout(() => {
            selectedItem.classList.add('active');
        }, index * 100);
    });
    
    // Show/hide the selected signature section
    const selectedSection = document.getElementById('selectedSignatureSection');
    if (checkboxes.length > 0) {
        selectedSection.style.display = 'block';
    } else {
        selectedSection.style.display = 'none';
    }
}

function createSelectedItem(text) {
    const selectedItem = document.createElement('div');
    selectedItem.className = 'selected-item';
    selectedItem.innerHTML = `
        <span class="selected-icon">✓</span>
        <span class="selected-text">${text}</span>
    `;
    return selectedItem;
}

// Export functions for use in other scripts if needed
window.signatureSelection = {
    updateSelectedDisplay: updateSelectedDisplay,
    updateCardState: updateCardState
};