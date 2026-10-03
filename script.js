// ======================================================
// PROJECT FORMULATION, BRAINSTORMING &
// AI-ASSISTED DEVELOPMENT TOOLS WORKSHOP
// CERTIFICATE DOWNLOAD PORTAL
// ======================================================

let students = [];


// ======================================================
// LOAD STUDENT DATA FROM EXCEL
// ======================================================

async function loadStudentData() {

    try {

        const response = await fetch(
            "Attendance_Project_Workshop.xlsx"
        );

        if (!response.ok) {
            throw new Error(
                "Excel file could not be loaded."
            );
        }

        const arrayBuffer =
            await response.arrayBuffer();

        const workbook = XLSX.read(
            arrayBuffer,
            {
                type: "array"
            }
        );


        // --------------------------------------------------
        // READ SHEET2
        // --------------------------------------------------

        const worksheet =
            workbook.Sheets["Sheet2"];

        if (!worksheet) {
            throw new Error(
                "Sheet2 was not found in the Excel file."
            );
        }


        const data =
            XLSX.utils.sheet_to_json(
                worksheet,
                {
                    defval: ""
                }
            );


        // --------------------------------------------------
        // CREATE STUDENT RECORDS
        // --------------------------------------------------

        students = data
            .filter(row => row["S.No"])
            .map(row => {

                const sno =
                    Number(row["S.No"]);

                const rollNumber =
                    String(
                        row["Roll Number"] || ""
                    ).trim();

                const name =
                    String(
                        row["Participant's Full Name"] || ""
                    ).trim();

                const course =
                    String(
                        row["Course"] || ""
                    ).trim();


                // --------------------------------------------------
                // PDF FILE NAME
                //
                // Example:
                // Dev Bhati
                // becomes
                // Dev Bhati.pdf
                // --------------------------------------------------

                const certificateFileName =
                    name + ".pdf";


                return {

                    sno: sno,

                    rollNumber: rollNumber,

                    name: name,

                    course: course,

                    certificateFile:
                        "certificates/" +
                        certificateFileName
                };

            });


        console.log(
            "Students loaded:",
            students.length
        );

        console.log(
            "Student records:",
            students
        );

    }


    catch (error) {

        console.error(
            "Error loading student data:",
            error
        );


        document.getElementById(
            "result"
        ).innerHTML = `

            <div class="not-found">

                <h3>
                    ⚠️ Unable to Load Certificate Data
                </h3>

                <p>
                    The certificate database
                    could not be loaded.
                </p>

                <p>
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
        document.getElementById(
            "searchInput"
        );

    const result =
        document.getElementById(
            "result"
        );


    const keyword =
        searchInput.value
            .trim()
            .toLowerCase();


    // --------------------------------------------------
    // EMPTY SEARCH
    // --------------------------------------------------

    if (keyword === "") {

        result.innerHTML = `

            <h3>
                Enter Search Details
            </h3>

            <p>
                Please enter your Name
                or Roll Number.
            </p>

        `;

        return;
    }


    // --------------------------------------------------
    // SEARCH STUDENTS
    // --------------------------------------------------

    const matches =
        students.filter(student => {

            const name =
                student.name.toLowerCase();

            const rollNumber =
                student.rollNumber.toLowerCase();


            return (
                name.includes(keyword) ||
                rollNumber === keyword
            );

        });


    // --------------------------------------------------
    // NO MATCH FOUND
    // --------------------------------------------------

    if (matches.length === 0) {

        result.innerHTML = `

            <div class="not-found">

                <h3>
                    ❌ Certificate Not Found
                </h3>

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


    // --------------------------------------------------
    // ONE STUDENT FOUND
    // --------------------------------------------------

    if (matches.length === 1) {

        displayCertificate(
            matches[0]
        );

        return;
    }


    // --------------------------------------------------
    // MULTIPLE STUDENTS FOUND
    // --------------------------------------------------

    result.innerHTML = `

        <div class="multiple-results">

            <h3>
                ✓ Multiple Certificates Found
            </h3>

            <p>
                More than one student matches
                this name.
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
                                Project Formulation,
                                Brainstorming &
                                AI-Assisted Development
                                Tools Workshop
                            </span>

                        </div>


                        <div class="detail-row">

                            <span class="label">
                                Date
                            </span>

                            <span>
                                5 September, 2026
                            </span>

                        </div>

                    </div>


                    <a
                        href="${encodeURI(student.certificateFile)}"
                        class="download-btn"
                        download
                    >
                        Download Certificate
                    </a>

                </div>

            `).join("")}

        </div>

    `;

}


// ======================================================
// DISPLAY ONE CERTIFICATE
// ======================================================

function displayCertificate(student) {

    const result =
        document.getElementById(
            "result"
        );


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
                        Project Formulation,
                        Brainstorming &
                        AI-Assisted Development
                        Tools Workshop
                    </span>

                </div>


                <div class="detail-row">

                    <span class="label">
                        Date
                    </span>

                    <span>
                        5 September, 2026
                    </span>

                </div>


            </div>


            <!-- DOWNLOAD BUTTON -->

            <a
                href="${encodeURI(student.certificateFile)}"
                class="download-btn"
                download
            >
                Download Certificate
            </a>


        </div>

    `;

}


// ======================================================
// PRESS ENTER TO SEARCH
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
// CLEAR RESULT WHEN SEARCH BOX IS EMPTY
// ======================================================

document
    .getElementById("searchInput")
    .addEventListener(
        "input",
        function() {

            if (
                this.value.trim() === ""
            ) {

                document.getElementById(
                    "result"
                ).innerHTML = `

                    <h3>
                        Welcome 👋
                    </h3>

                    <p>
                        Enter your Name or Roll Number
                        to verify and download
                        your certificate.
                    </p>

                `;

            }

        }
    );


// ======================================================
// LOAD DATA WHEN PAGE OPENS
// ======================================================

loadStudentData();
