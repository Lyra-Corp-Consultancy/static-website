document.getElementById("exportBtn").addEventListener("click", function() {
    fetchMicrosoftGraphDataAndExport();
});

function fetchMicrosoftGraphDataAndExport() {
    const token = "eyJ0eXAiOiJKV1QiLCJub25jZSI6IjgxQzVJX3lleFZ6TTh0Q3FRN3o1bXZsOUlGMU9fUzdCbjNzMkplUUNJdlEiLCJhbGciOiJSUzI1NiIsIng1dCI6IkwxS2ZLRklfam5YYndXYzIyeFp4dzFzVUhIMCIsImtpZCI6IkwxS2ZLRklfam5YYndXYzIyeFp4dzFzVUhIMCJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9jNTE1NjhjMi00OTA4LTRlNDgtODlhZS0zMmQzM2EyYzM2MjQvIiwiaWF0IjoxNzE1MzE1NjA2LCJuYmYiOjE3MTUzMTU2MDYsImV4cCI6MTcxNTQwMjMwNywiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFUUUF5LzhXQUFBQVRjUkM0Q3VDYzBkeHRmRU9SdSsxekx1V3g0dmdnSjNJdXdoWVZ5bFpWTFAvdFZvenZKVi9FZDdmUkRNYlppaXYiLCJhbXIiOlsicHdkIiwicnNhIl0sImFwcF9kaXNwbGF5bmFtZSI6IkdyYXBoIEV4cGxvcmVyIiwiYXBwaWQiOiJkZThiYzhiNS1kOWY5LTQ4YjEtYThhZC1iNzQ4ZGE3MjUwNjQiLCJhcHBpZGFjciI6IjAiLCJkZXZpY2VpZCI6ImIxYzljMTFlLTg0OTMtNGYzNS04NjQyLTZkMzM4YmE1MjNiZCIsImZhbWlseV9uYW1lIjoiQWhhbW1lZEFsaSIsImdpdmVuX25hbWUiOiJGYXJpcyIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjEwMy4xMzAuOTEuMjUyIiwibmFtZSI6IkZhcmlzIEFoYW1tZWRBbGkiLCJvaWQiOiI3NzdmZjBiZS1kYzk2LTQ0YWEtODk5My1iNTYwODZiMmZkYmEiLCJwbGF0ZiI6IjMiLCJwdWlkIjoiMTAwMzIwMDM3MUVEQzQ1QiIsInJoIjoiMC5BVW9Bd21nVnhRaEpTRTZKcmpMVE9pdzJKQU1BQUFBQUFBQUF3QUFBQUFBQUFBQ0pBTVEuIiwic2NwIjoib3BlbmlkIHByb2ZpbGUgVXNlci5SZWFkIGVtYWlsIiwic2lnbmluX3N0YXRlIjpbImttc2kiXSwic3ViIjoib1ZGUnRVOXNpLWp0SmdlNjFnWUstdnFqdFMwVUNyTmMzMjVPdkVYb09fUSIsInRlbmFudF9yZWdpb25fc2NvcGUiOiJBUyIsInRpZCI6ImM1MTU2OGMyLTQ5MDgtNGU0OC04OWFlLTMyZDMzYTJjMzYyNCIsInVuaXF1ZV9uYW1lIjoiRmFyaXNBaGFtbWVkQWxpQGx5cmFjb3JwLmluIiwidXBuIjoiRmFyaXNBaGFtbWVkQWxpQGx5cmFjb3JwLmluIiwidXRpIjoiZTFob1c1bXJNa1NvZXVwLWFNT0dBQSIsInZlciI6IjEuMCIsIndpZHMiOlsiYjc5ZmJmNGQtM2VmOS00Njg5LTgxNDMtNzZiMTk0ZTg1NTA5Il0sInhtc19jYyI6WyJDUDEiXSwieG1zX3NzbSI6IjEiLCJ4bXNfc3QiOnsic3ViIjoienljSmFaNTBjT1J0YlhGakZxZ3B0OEV1eng4NHk3S01VeW1tcVM3aU5DWSJ9LCJ4bXNfdGNkdCI6MTY5MTU4OTIzNH0.cTfg3jTGpUFbgJTocwIiumkpgGCdGDe1B8eDkkeEQnFdNcndW5sYCyKgy32k2BKEzIgS1nRyz3jvl6UgavIOfhVBMv7E_mYK3i7Oh4fOH4LDOaoZ0BSCuclMiSjdusZjaiBI44zI4GTYuZOVIKPYVKsYB5fRWty5bSXAJowe1qBy-rU0te5RXSW9vgzXw1a-_kzpudt9aHJUUZ3EeTlGFYcBNfZHxn5Gq02lmkXuqCMnqWmNxcQFvRbaN6ukaFz8F7kHrE3fms-xFIC5IJJPkug03zm7ZYmEIZyYT9ROoR-PLbvvnHm98ZmsljhrTuy-pfDxjsB4S5eYwLMbNt521Q";
    const headers = {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json"
    };

    const apiEndpoint = "https://graph.microsoft.com/v1.0/users";

    fetch(apiEndpoint, { headers }) 
        .then(response => response.json())
        .then(data => {
            // Prepare data for Excel
            const usersData = data.value.map(user => [user.displayName, user.mail, user.department]);

            // Create a new Excel workbook
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet([["Name", "Email", "Department"], ...usersData]);

            // Add worksheet to the workbook
            XLSX.utils.book_append_sheet(wb, ws, "Users");

            // Generate Excel file
            const wbout = XLSX.write(wb, { bookType: "xlsx", type: "binary" });

            // Trigger file download
            saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), "user_data.xlsx");
        })
        .catch(error => {
            console.error("Error:", error);
        });
}

// Function to convert string to array buffer
function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
}
