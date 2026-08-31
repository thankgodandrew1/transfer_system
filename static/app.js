const form = document.querySelector('[data-generation-form], #generation-form');
const overlay = document.querySelector('#processing-overlay');
const processingMessage = document.querySelector('#processing-message');

document.querySelectorAll('input[type="file"]').forEach((input) => {
  input.addEventListener('change', () => {
    const field = input.closest('label');
    const label = field?.querySelector('[data-file-label]');
    if (label && input.files?.length) {
      label.textContent = input.files[0].name;
      field.classList.add('has-file');
    }
  });
});

if (form && overlay) {
  form.addEventListener('submit', () => {
    overlay.hidden = false;
    document.body.classList.add('is-processing');
    const movementMessages = [
      'Reading both Transfer News files…',
      'Matching missionaries across transfer cycles…',
      'Mapping each area to its apartment…',
      'Applying manual movement exceptions…',
      'Rendering the Word plan and review files…',
    ];
    const newsMessages = [
      'Validating uploaded records…',
      'Matching missionaries and previous assignments…',
      'Applying ground-truth verification…',
      'Rendering Word and PDF documents…',
      'Preparing your secure download package…',
    ];
    const messages = form.dataset.processingKind === 'movement' ? movementMessages : newsMessages;
    let index = 0;
    window.setInterval(() => {
      index = Math.min(index + 1, messages.length - 1);
      if (processingMessage) processingMessage.textContent = messages[index];
    }, 4500);
  });
}
