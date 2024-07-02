document.getElementById('myForm').addEventListener('submit', function(e) {
  e.preventDefault();
  
  var formData = new FormData(); // Create a new FormData object
  
  // Add form input values to the FormData object
  formData.append('nameInput', document.getElementById('nameInput').value);
  formData.append('batchInput', document.getElementById('batchInput').value);
  formData.append('resourceTypeInput', document.getElementById('resourceTypeInput').value);
  formData.append('courseNoInput', document.getElementById('courseNoInput').value);
  formData.append('teacherNameInput', document.getElementById('teacherNameInput').value);
  formData.append('fileInput', document.getElementById('fileInput').files[0]);
  
  var xhr = new XMLHttpRequest();
  xhr.open('POST', '/upload', true);
  xhr.onload = function() {
    if (xhr.status === 200) {
      console.log('File uploaded successfully.');
      // Redirect to the success page
      window.location.href = '/success.html';
    } else {
      console.error('Error uploading file.');
      // Handle error response
    }
  };
  xhr.send(formData);
});

// Function to show modal
function showModal() {
  const modal = document.getElementById("myModal");
  modal.style.display = "block";

  // Close modal when clicking on the close button (×)
  const closeButton = document.getElementsByClassName("close")[0];
  closeButton.onclick = function () {
    modal.style.display = "none";
  };

  // Close modal when clicking outside the modal content
  window.onclick = function (event) {
    if (event.target == modal) {
      modal.style.display = "none";
    }
  };
}

// Update your existing upload event listener to show the modal
document.getElementById("myForm").addEventListener("submit", function (event) {
  event.preventDefault(); // Prevent form submission to another page
  showProgressBar(function () {
    showModal(); // Show modal after progress bar finishes
  });
});

