document.getElementById("idForm").addEventListener("submit", async function (e) {
    e.preventDefault();

    const name = document.getElementById("name").value;
    const email = document.getElementById("email").value;
    const idNumber = document.getElementById("idNumber").value;

    // Display ID Card on Frontend
    document.getElementById("cardName").textContent = name;
    document.getElementById("cardEmail").textContent = email;
    document.getElementById("cardIdNumber").textContent = idNumber;
    document.getElementById("idCard").classList.remove("hidden");

    // Send Data to Backend
    await fetch("https://id-1-dxcx.onrender.com", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name, email, idNumber })
    });
    
});
