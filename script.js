// ======================================================
// LOAD ACTUAL STUDENT DATA FROM EXCEL
// ======================================================

let students = [];


// ======================================================
// LOAD EXCEL FILE
// ======================================================

async function loadStudentData() {

    try {

        const response = await fetch("Attendance_Git_GitHub.xlsx");

        if (!response.ok) {
            throw new Error("Excel file could not be loaded.");
        }

        const arrayBuffer = await response.arrayBuffer();

        const workbook = XLSX.read(arrayBuffer, {
            type: "array"
        });

        // Use Sheet2 because the actual participant data
        // is stored there
        const worksheet = workbook.Sheets["Sheet2"];

        if (!worksheet) {
            throw new Error("Sheet2 was not found in the Excel file.");
        }

        const data = XLSX.utils.sheet_to_json(worksheet, {
            defval: ""
        });


        // --------------------------------------------------
        // Create student records
        // --------------------------------------------------

        const usedFileNames = {};

        students = data
            .filter(row => row["Sno"])
            .map(row => {

                const sno = Number(row["Sno"]);

                const rollNumber =
                    String(row["Roll Number  "] || "").trim();

                const course =
                    String(row["Course  "] || "").trim();

                const name =
                    String(row["Participant's Full Name "] || "").trim();


                // ------------------------------------------
                // Create PDF filename
                // ------------------------------------------

               let certificateFileName = name + ".pdf";


                // Special handling for duplicate Farhan Khan
                if (name.toLowerCase() === "farhan khan") {
                
                    if (course.toLowerCase().includes("iii")) {
                
                        certificateFileName = "Farhan Khan_BCA_III.pdf";
                
                    } else if (course.toLowerCase().includes("v")) {
                
                        certificateFileName = "Farhan Khan.pdf";
                
                    }
                
                }

                usedFileNames[certificateFileName] = true;


                return {

                    sno: sno,

                    rollNumber: rollNumber,

                    course: course,

                    name: name,

                    certificateFile:
                        "certificates/" +
                        certificateFileName

                };

            });


        console.log(
            "Student records loaded:",
            students.length
        );

    }

    catch (error) {

        console.error(
            "Error loading student data:",
            error
        );

        document.getElementById("result").innerHTML = `

            <div class="not-found">

                <h3>⚠️ Unable to Load Certificate Data</h3>

                <p>
                    The certificate database could not be loaded.
                    Please try again later.
                </p>

            </div>

        `;

    }

}


// ======================================================
// SEARCH CERTIFICATE
// ======================================================

function searchCertificate() {

    const searchInput =
        document.getElementById("searchInput");

    const result =
        document.getElementById("result");

    const keyword =
        searchInput.value
            .trim()
            .toLowerCase();


    // Check if search box is empty

    if (keyword === "") {

        result.innerHTML = `

            <div class="message-box">

                <h3>Enter Search Details</h3>

                <p>
                    Please enter your Name or Roll Number.
                </p>

            </div>

        `;

        return;

    }


    // Find ALL matching students
    const matches = students.filter(student => {

        const name =
            student.name.toLowerCase();

        const rollNumber =
            student.rollNumber.toLowerCase();


        return (
            name.includes(keyword) ||
            rollNumber === keyword
        );

    });


    // No certificate found

    if (matches.length === 0) {

        result.innerHTML = `

            <div class="not-found">

                <h3>❌ Certificate Not Found</h3>

                <p>
                    We could not find a certificate
                    matching the information entered.
                </p>

                <p>
                    Please check the student's
                    Name or Roll Number and try again.
                </p>

            </div>

        `;

        return;

    }


    // Display all matching certificates

    if (matches.length === 1) {

        displayCertificate(matches[0]);

        return;

    }


    // Multiple students found

    result.innerHTML = `

        <div class="multiple-results">

            <h3>✓ Multiple Certificates Found</h3>

            <p>
                More than one student matches this name.
                Please select the correct certificate.
            </p>

            ${matches.map(student => `

                <div class="certificate-result">

                    <div class="verified-badge">
                        ✓ Verified Certificate
                    </div>

                    <h2>
                        ${student.name}
                    </h2>

                    <div class="student-details">

                        <div class="detail-row">

                            <span class="label">
                                Roll Number
                            </span>

                            <span>
                                ${student.rollNumber}
                            </span>

                        </div>

                        <div class="detail-row">

                            <span class="label">
                                Course
                            </span>

                            <span>
                                ${student.course}
                            </span>

                        </div>

                        <div class="detail-row">

                            <span class="label">
                                Workshop
                            </span>

                            <span>
                                Git & GitHub Workshop
                            </span>

                        </div>

                        <div class="detail-row">

                            <span class="label">
                                Date
                            </span>

                            <span>
                                18 July, 2026
                            </span>

                        </div>

                    </div>

                   
                    <a
                        href="${student.certificateFile}"
                        class="download-btn"
                        onclick="downloadCertificate(event, '${student.certificateFile}', '${student.name}')"
                    >
                        Download Certificate
                    </a>

                </div>

            `).join("")}

        </div>

    `;

}

// ======================================================
// DISPLAY CERTIFICATE DETAILS
// ======================================================

function displayCertificate(student) {

    const result =
        document.getElementById("result");


    result.innerHTML = `

        <div class="certificate-result">

            <div class="verified-badge">

                ✓ Verified Certificate

            </div>


            <h2>
                ${student.name}
            </h2>


            <div class="student-details">


                <div class="detail-row">

                    <span class="label">
                        Roll Number
                    </span>

                    <span>
                        ${student.rollNumber}
                    </span>

                </div>


                <div class="detail-row">

                    <span class="label">
                        Course
                    </span>

                    <span>
                        ${student.course}
                    </span>

                </div>


                <div class="detail-row">

                    <span class="label">
                        Workshop
                    </span>

                    <span>
                        Git & GitHub Workshop
                    </span>

                </div>


                <div class="detail-row">

                    <span class="label">
                        Date
                    </span>

                    <span>
                        18 July, 2026
                    </span>

                </div>


            </div>


           
            <a
                href="${student.certificateFile}"
                class="download-btn"
                onclick="downloadCertificate(event, '${student.certificateFile}', '${student.name}')"
            >
                Download Certificate
            </a>
        </div>

    `;

}


// ======================================================
// SEARCH WHEN USER PRESSES ENTER
// ======================================================

document
    .getElementById("searchInput")
    .addEventListener(
        "keypress",
        function(event) {

            if (event.key === "Enter") {

                searchCertificate();

            }

        }
    );


// ======================================================
// CLEAR OLD RESULT WHEN USER STARTS NEW SEARCH
// ======================================================

document
    .getElementById("searchInput")
    .addEventListener(
        "input",
        function() {

            if (this.value.trim() === "") {

                document.getElementById("result").innerHTML = `

                    <h3>Welcome 👋</h3>

                    <p>
                        Enter your Name or Roll Number
                        to verify and download your certificate.
                    </p>

                `;

            }

        }
    );


async function downloadCertificate(event, filePath, studentName) {
    event.preventDefault();

    try {
        const response = await fetch(filePath);

        if (!response.ok) {
            throw new Error("Certificate could not be downloaded.");
        }

        const blob = await response.blob();

        const url = window.URL.createObjectURL(blob);

        const link = document.createElement("a");
        link.href = url;
        link.download = `${studentName} - Git & GitHub Workshop Certificate.pdf`;

        document.body.appendChild(link);
        link.click();

        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);

    } catch (error) {
        console.error("Download error:", error);
        alert("Unable to download the certificate. Please try again.");
    }
}


// ======================================================
// LOAD DATA WHEN PAGE OPENS
// ======================================================

loadStudentData();
