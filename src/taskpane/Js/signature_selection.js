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
    const checkboxes = document.querySelectorAll('.option-card input[type="checkbox"]:checked, input[type="checkbox"]:checked');
    
    if (!selectedOptionsContainer) return;
    
    // Clear existing selected items
    selectedOptionsContainer.innerHTML = '';
    
    // Add selected items with animation
    checkboxes.forEach((checkbox, index) => {
        let label;
        const optionCard = checkbox.closest('.option-card');
        if (optionCard) {
            label = optionCard.querySelector('.option-title').textContent;
        } else {
            // Fallback for checkboxes not in option-card structure
            const labelElement = document.querySelector(`label[for="${checkbox.id}"]`);
            label = labelElement ? labelElement.textContent : checkbox.id.replace('Signature', '').replace(/([A-Z])/g, ' $1').trim();
        }
        
        const selectedItem = createSelectedItem(label);
        selectedOptionsContainer.appendChild(selectedItem);
        
        // Trigger animation after a small delay
        setTimeout(() => {
            selectedItem.classList.add('active');
        }, index * 100);
    });
    
    // Show/hide the selected signature section (try both possible element IDs)
    let selectedSection = document.getElementById('selectedSignatureSection');
    if (!selectedSection) {
        selectedSection = document.getElementById('selectedConfigurationSection');
    }
    
    if (selectedSection) {
        if (checkboxes.length > 0) {
            selectedSection.style.display = 'block';
        } else {
            selectedSection.style.display = 'none';
        }
    }
}

function createSelectedItem(text) {
    const selectedItem = document.createElement('div');
    selectedItem.className = 'selected-item';
    selectedItem.style.cssText = "display: inline-flex; align-items: center; background-color: #0078d4; color: white; padding: 4px 8px; border-radius: 12px; font-size: 12px; font-weight: 500;";
    selectedItem.innerHTML = `<span style="margin-right: 4px;">✓</span>${text}`;
    return selectedItem;
}

// Export functions for use in other scripts if needed
window.signatureSelection = {
    updateSelectedDisplay: updateSelectedDisplay,
    updateCardState: updateCardState
};