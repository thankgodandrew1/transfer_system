const form = document.querySelector('#generation-form');
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
    const messages = [
      'Validating uploaded records…',
      'Matching missionaries and previous assignments…',
      'Applying ground-truth verification…',
      'Rendering Word and PDF documents…',
      'Preparing your secure download package…',
    ];
    let index = 0;
    window.setInterval(() => {
      index = Math.min(index + 1, messages.length - 1);
      if (processingMessage) processingMessage.textContent = messages[index];
    }, 4500);
  });
}
