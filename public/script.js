const msalInstance = new msal.PublicClientApplication({
    auth: {
        clientId: "<client-id-goes-here>",
        authority: "https://login.microsoftonline.com/<tenant-id-goes-here>",
        redirectUri: "http://localhost:8000",
    },
});

let allUsers = [];
let allDepartments = [];


async function login() {
    try {
        const loginResponse = await msalInstance.loginPopup({
            scopes: ["User.Read.All", "Directory.Read.All", "Mail.Send"],
        });
        msalInstance.setActiveAccount(loginResponse.account);
        alert("Login successful.");
        await fetchUsers();
        await fetchDepartments();
        
    } catch (error) {
        console.error("Login failed:", error);
        alert("Login failed.");
    }
}




function logout() {
    msalInstance.logoutPopup().then(() => alert("Logout successful."));
}

// Fetch Recently Created Users
async function fetchUsers() {
    const currentDate = new Date();
    const startDate = new Date(currentDate.setDate(currentDate.getDate() - 180)).toISOString();

    const response = await callGraphApi(`/users?$filter=createdDateTime ge ${startDate}&$select=displayName,userPrincipalName,mail,department,assignedLicenses,createdDateTime`);
    allUsers = response.value;
}

// Fetch Departments
async function fetchDepartments() {
    const departments = [...new Set(allUsers.map(user => user.department).filter(Boolean))];
    populateDropdown("departmentDropdown", departments.map(dep => ({ id: dep, name: dep })));
}



// Populate Dropdown
function populateDropdown(dropdownId, items) {
    const dropdown = document.getElementById(dropdownId);
    dropdown.innerHTML = `<option value="">Select</option>`;
    items.forEach(item => {
        const option = document.createElement("option");
        option.value = item.id;
        option.textContent = item.name;
        dropdown.appendChild(option);
    });
}




// search functionality

function search() {
    const searchText = document.getElementById("searchBox").value.toLowerCase();
    const fromDate = document.getElementById("fromDate").value ? new Date(document.getElementById("fromDate").value).toISOString() : null;
    const toDate = document.getElementById("toDate").value ? new Date(document.getElementById("toDate").value).toISOString() : null;
    const licenseStatus = document.getElementById("licenseDropdown").value;
    const department = document.getElementById("departmentDropdown").value;
    

    const filteredUsers = allUsers.filter(user => {
        const matchesSearchText = searchText
            ? (user.displayName?.toLowerCase().includes(searchText) ||
               user.userPrincipalName?.toLowerCase().includes(searchText) ||
               user.mail?.toLowerCase().includes(searchText))
            : true;

        const matchesDateRange = fromDate && toDate
            ? new Date(user.createdDateTime) >= new Date(fromDate) && new Date(user.createdDateTime) <= new Date(toDate)
            : true;

        const matchesLicense = licenseStatus
            ? (licenseStatus === "Licensed" && user.assignedLicenses.length > 0) ||
              (licenseStatus === "Unlicensed" && user.assignedLicenses.length === 0)
            : true;

        const matchesDepartment = department
            ? user.department === department
            : true;

        

        return matchesSearchText && matchesDateRange && matchesLicense && matchesDepartment;
    });

    if (filteredUsers.length === 0) {
        alert("No matching results found.");
    }

    displayResults(filteredUsers);
}



// Display Results
function displayResults(users) {
    const outputBody = document.getElementById("outputBody");
    outputBody.innerHTML = users.map(user => `
        <tr>
            <td>${user.displayName || "N/A"}</td>
            <td>${user.userPrincipalName || "N/A"}</td>
            <td>${user.mail || "N/A"}</td>
            <td>${user.department || "N/A"}</td>
            <td>${user.role || "N/A"}</td>
            <td>${user.assignedLicenses.length > 0 ? "Licensed" : "Unlicensed"}</td>
            <td>${new Date(user.createdDateTime).toLocaleDateString()}</td>
        </tr>
    `).join("");
}

// Utility Functions

async function callGraphApi(endpoint, method = "GET", body = null) {
    const account = msalInstance.getActiveAccount();
    if (!account) throw new Error("Please log in first.");

    try {
        const tokenResponse = await msalInstance.acquireTokenSilent({
            scopes: ["User.ReadWrite.All", "Directory.ReadWrite.All", "Mail.Send"],
            account,
        });

        const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
            method,
            headers: {
                Authorization: `Bearer ${tokenResponse.accessToken}`,
                "Content-Type": "application/json",
            },
            body: body ? JSON.stringify(body) : null, // Ensure the body is serialized if provided
        });

        if (response.ok) {
            const contentType = response.headers.get("content-type");
            if (contentType && contentType.includes("application/json")) {
                return await response.json(); // Parse JSON response
            }
            return {}; // Handle responses with no body
        } else {
            const errorText = await response.text();
            console.error(`Graph API Error (${response.status}):`, errorText);
            throw new Error(`Graph API call failed: ${response.status} ${response.statusText}`);
        }
    } catch (error) {
        console.error("Error in callGraphApi:", error);
        throw error;
    }
}





// Download Report as CSV
function downloadReportAsCSV() {
    const headers = ["Display Name", "UPN", "Email", "Department", "Role", "License Status", "Created Date"];
    const rows = [...document.querySelectorAll("#outputBody tr")].map(tr =>
        [...tr.querySelectorAll("td")].map(td => td.textContent)
    );

    if (!rows.length) {
        alert("No data available to download.");
        return;
    }

    const csvContent = [headers.join(","), ...rows.map(row => row.join(","))].join("\n");
    const blob = new Blob([csvContent], { type: "text/csv" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "Recently_Created_Users_Report.csv";
    link.click();
}

// Mail Report to Admin
async function sendReportAsMail() {
    const adminEmail = document.getElementById("adminEmail").value;

    if (!adminEmail) {
        alert("Please provide an admin email.");
        return;
    }

    const headers = [...document.querySelectorAll("#outputHeader th")].map(th => th.textContent);
    const rows = [...document.querySelectorAll("#outputBody tr")].map(tr =>
        [...tr.querySelectorAll("td")].map(td => td.textContent)
    );

    if (!rows.length) {
        alert("No data to send via email.");
        return;
    }

    const emailContent = rows.map(row => `<tr>${row.map(cell => `<td>${cell}</td>`).join("")}</tr>`).join("");
    const emailBody = `
        <table border="1">
            <thead>
                <tr>${headers.map(header => `<th>${header}</th>`).join("")}</tr>
            </thead>
            <tbody>${emailContent}</tbody>
        </table>
    `;

    const message = {
        message: {
            subject: "Recently Created Users Report",
            body: { contentType: "HTML", content: emailBody },
            toRecipients: [{ emailAddress: { address: adminEmail } }],
        },
    };

    try {
        await callGraphApi("/me/sendMail", "POST", message);
        alert("Report sent successfully!");
    } catch (error) {
        console.error("Error sending report:", error);
        alert("Failed to send the report.");
    }
}
