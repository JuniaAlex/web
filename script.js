var gk_isXlsx = false;
        var gk_xlsxFileLookup = {};
        var gk_fileData = {};
        function filledCell(cell) {
          return cell !== '' && cell != null;
        }
        function loadFileData(filename) {
        if (gk_isXlsx && gk_xlsxFileLookup[filename]) {
            try {
                var workbook = XLSX.read(gk_fileData[filename], { type: 'base64' });
                var firstSheetName = workbook.SheetNames[0];
                var worksheet = workbook.Sheets[firstSheetName];

                // Convert sheet to JSON to filter blank rows
                var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: false, defval: '' });
                // Filter out blank rows (rows where all cells are empty, null, or undefined)
                var filteredData = jsonData.filter(row => row.some(filledCell));

                // Heuristic to find the header row by ignoring rows with fewer filled cells than the next row
                var headerRowIndex = filteredData.findIndex((row, index) =>
                  row.filter(filledCell).length >= filteredData[index + 1]?.filter(filledCell).length
                );
                // Fallback
                if (headerRowIndex === -1 || headerRowIndex > 25) {
                  headerRowIndex = 0;
                }

                // Convert filtered JSON back to CSV
                var csv = XLSX.utils.aoa_to_sheet(filteredData.slice(headerRowIndex)); // Create a new sheet from filtered array of arrays
                csv = XLSX.utils.sheet_to_csv(csv, { header: 1 });
                return csv;
            } catch (e) {
                console.error(e);
                return "";
            }
        }
        return gk_fileData[filename] || "";
        }
        


        // Data
        let subjects = [];
        let tasks = [];
        let timer;
        let timeLeft = 25 * 60 * 60; // Default to 25 minutes (in seconds)
        let initialTimeLeft = 25 * 60 * 60; // Store the initial duration for resetting

        const quotes = [
            "Focus on progress, not perfection.",
            "Every study session brings you closer to your goals.",
            "Stay curious, keep learning, and never give up!",
            "Small steps today lead to big victories tomorrow.",
            "Your effort now shapes your future success."
        ];

        // Load data from localStorage
        function loadData() {
            console.log("loadData called");
            const savedSubjects = localStorage.getItem('subjects');
            const savedTasks = localStorage.getItem('tasks');
            if (savedSubjects) subjects = JSON.parse(savedSubjects);
            if (savedTasks) tasks = JSON.parse(savedTasks).map(task => ({
                ...task,
                deadline: task.deadline ? new Date(task.deadline) : null
            }));
        }

        // Save data to localStorage
        function saveData() {
            console.log("saveData called");
            localStorage.setItem('subjects', JSON.stringify(subjects));
            localStorage.setItem('tasks', JSON.stringify(tasks));
        }

        // Calculate average marks
        function calculateAverageMarks() {
            if (subjects.length === 0) return 0;
            const total = subjects.reduce((sum, subject) => sum + subject.grade, 0);
            return Math.round(total / subjects.length);
        }

        // Update Overall Score
        function updateOverallScore() {
            console.log("updateOverallScore called");
            const average = calculateAverageMarks();
            const circumference = 2 * Math.PI * 50; // Radius = 50
            const progress = (average / 100) * circumference;
            const progressCircle = document.getElementById('progress-circle');
            const scoreText = document.getElementById('overall-score-text');
            if (progressCircle && scoreText) {
                progressCircle.setAttribute('stroke-dasharray', `${progress} ${circumference}`);
                scoreText.textContent = `${average}%`;
            }
        }

        // Display random quote
        function displayQuote() {
            console.log("displayQuote called");
            const quoteElement = document.getElementById("motivationalQuote");
            if (quoteElement) {
                quoteElement.textContent = quotes[Math.floor(Math.random() * quotes.length)];
            }
        }

        // Show section
        function showSection(sectionId) {
            console.log(`showSection called with sectionId: ${sectionId}`);
            document.querySelectorAll(".section").forEach(section => section.classList.remove("active"));
            document.querySelectorAll(".sidebar a").forEach(a => a.classList.remove("active"));
            const section = document.getElementById(sectionId);
            if (section) section.classList.add("active");
            const sidebarLink = document.querySelector(`.sidebar a[onclick="showSection('${sectionId}')"]`);
            if (sidebarLink) sidebarLink.classList.add("active");
            if (sectionId === "subjects") displaySubjects();
            if (sectionId === "todo") displayTasks();
            if (sectionId === "calendar") generateCalendar();
            if (sectionId === "home") {
                displayQuote();
                updateOverallScore();
            }
            if (sectionId === "timer") {
                updateTimerDisplay();
            }
            // Force style refresh for dark mode
            document.body.classList.toggle("dark-mode", document.body.classList.contains("dark-mode"));
        }

        // Mode Toggle
        function toggleMode() {
            console.log("toggleMode called");
            const isDarkMode = document.body.classList.toggle("dark-mode");
            const modeToggle = document.getElementById("modeToggle");
            if (modeToggle) {
                modeToggle.textContent = isDarkMode ? "Light" : "Dark";
            }
            // Force repaint to ensure styles apply
            document.body.style.display = 'none';
            document.body.offsetHeight; // Trigger reflow
            document.body.style.display = '';
        }

        // Subjects Management
        function addSubject() {
            console.log("addSubject called");
            const nameInput = document.getElementById("subjectName");
            const gradeInput = document.getElementById("subjectGrade");
            const name = nameInput ? nameInput.value.trim() : "";
            const grade = gradeInput ? parseInt(gradeInput.value) : NaN;
            if (!name) {
                alert("Please enter a subject name.");
                return;
            }
            if (isNaN(grade) || grade < 0 || grade > 100) {
                alert("Please enter a valid grade (0-100).");
                return;
            }
            subjects.push({ name, grade });
            saveData();
            displaySubjects();
            if (nameInput) nameInput.value = "";
            if (gradeInput) nameInput.value = "";
            updateOverallScore();
        }

        function editSubject(index) {
            console.log(`editSubject called with index: ${index}`);
            const newName = prompt("Enter new subject name:", subjects[index].name);
            const newGrade = parseInt(prompt("Enter new grade (0-100):", subjects[index].grade));
            if (newName && newName.trim() && !isNaN(newGrade) && newGrade >= 0 && newGrade <= 100) {
                subjects[index].name = newName.trim();
                subjects[index].grade = newGrade;
                saveData();
                displaySubjects();
                updateOverallScore();
            } else if (newName !== null) {
                alert("Invalid input. Subject name cannot be empty, and grade must be between 0 and 100.");
            }
        }

        function deleteSubject(index) {
            console.log(`deleteSubject called with index: ${index}`);
            if (confirm("Are you sure you want to delete this subject?")) {
                subjects.splice(index, 1);
                saveData();
                displaySubjects();
                updateOverallScore();
            }
        }

        function displaySubjects() {
            console.log("displaySubjects called");
            const subjectsList = document.getElementById("subjectList");
            if (subjectsList) {
                subjectsList.innerHTML = "";
                subjects.forEach((subject, index) => {
                    const li = document.createElement("li");
                    li.innerHTML = `
                        <span>${subject.name}: ${subject.grade}%</span>
                        <div class="input-group">
                            <button onclick="editSubject(${index})" aria-label="Edit Subject">Edit</button>
                            <button onclick="deleteSubject(${index})" aria-label="Delete Subject">Delete</button>
                        </div>
                    `;
                    subjectsList.appendChild(li);
                });
            }
        }

        // To-Do List Management
        function addTask() {
            console.log("addTask called");
            const taskInput = document.getElementById("taskInput");
            const deadlineInput = document.getElementById("taskDeadline");
            const taskText = taskInput ? taskInput.value.trim() : "";
            const deadline = deadlineInput ? deadlineInput.value : "";
            if (!taskText) {
                alert("Please enter a task.");
                return;
            }
            tasks.push({
                text: taskText,
                completed: false,
                deadline: deadline ? new Date(deadline) : null
            });
            saveData();
            displayTasks();
            if (taskInput) taskInput.value = "";
            if (deadlineInput) deadlineInput.value = "";
        }

        function displayTasks() {
            console.log("displayTasks called");
            const taskList = document.getElementById("taskList");
            if (taskList) {
                taskList.innerHTML = "";
                tasks.forEach((task, index) => {
                    const li = document.createElement("li");
                    li.innerHTML = `
                        <span style="text-decoration: ${task.completed ? 'line-through' : 'none'}">
                            ${task.text}${task.deadline ? ` (Due: ${task.deadline.toLocaleDateString()})` : ''}
                        </span>
                        <div>
                            <button onclick="toggleTask(${index})" aria-label="${task.completed ? 'Undo Task' : 'Complete Task'}">${task.completed ? 'Undo' : 'Done'}</button>
                            <button onclick="removeTask(${index})" aria-label="Remove Task">Remove</button>
                        </div>
                    `;
                    taskList.appendChild(li);
                });
            }
        }

        function toggleTask(index) {
            console.log(`toggleTask called with index: ${index}`);
            tasks[index].completed = !tasks[index].completed;
            saveData();
            displayTasks();
        }

        function removeTask(index) {
            console.log(`removeTask called with index: ${index}`);
            tasks.splice(index, 1);
            saveData();
            displayTasks();
        }

        // Dynamic Calendar
        function generateCalendar() {
            console.log("generateCalendar called");
            const today = new Date();
            const year = today.getFullYear();
            const month = today.getMonth();
            const firstDay = new Date(year, month, 1).getDay();
            const daysInMonth = new Date(year, month + 1, 0).getDate();
            const currentDay = today.getDate();
            const calendarTitle = document.getElementById("calendarTitle");
            if (calendarTitle) {
                calendarTitle.textContent = `Calendar - ${today.toLocaleString('default', { month: 'long' })} ${year}`;
            }
            const calendarGrid = document.getElementById("calendarGrid");
            if (calendarGrid) {
                calendarGrid.innerHTML = '<div>Sun</div><div>Mon</div><div>Tue</div><div>Wed</div><div>Thu</div><div>Fri</div><div>Sat</div>';
                for (let i = 0; i < firstDay; i++) {
                    calendarGrid.innerHTML += '<div></div>';
                }
                for (let day = 1; day <= daysInMonth; day++) {
                    const date = new Date(year, month, day);
                    const hasEvent = tasks.some(task => task.deadline && task.deadline.toDateString() === date.toDateString());
                    const isToday = day === currentDay ? 'today' : '';
                    const isEvent = hasEvent ? 'event' : '';
                    calendarGrid.innerHTML += `<div class="${isToday} ${isEvent}" onclick="showTasksForDate('${date.toISOString()}')">${day}</div>`;
                }
            }
        }

        function showTasksForDate(dateStr) {
            console.log(`showTasksForDate called with date: ${dateStr}`);
            const date = new Date(dateStr);
            const tasksForDate = tasks.filter(task => task.deadline && task.deadline.toDateString() === date.toDateString());
            const taskList = tasksForDate.map(task => task.text).join('\n') || 'No tasks for this date.';
            alert(`Tasks for ${date.toLocaleDateString()}:\n${taskList}`);
        }

        // Pomodoro Timer
        function setCustomTimer() {
            console.log("setCustomTimer called");
            const hoursInput = document.getElementById("timerHours");
            const minutesInput = document.getElementById("timerMinutes");
            const secondsInput = document.getElementById("timerSeconds");
            const hours = hoursInput ? parseInt(hoursInput.value) || 0 : 0;
            const minutes = minutesInput ? parseInt(minutesInput.value) || 0 : 0;
            const seconds = secondsInput ? parseInt(secondsInput.value) || 0 : 0;

            // Validate input
            if (hours < 0 || hours > 23) {
                alert("Hours must be between 0 and 23.");
                return;
            }
            if (minutes < 0 || minutes > 59) {
                alert("Minutes must be between 0 and 59.");
                return;
            }
            if (seconds < 0 || seconds > 59) {
                alert("Seconds must be between 0 and 59.");
                return;
            }
            if (hours === 0 && minutes === 0 && seconds === 0) {
                alert("Please set a valid timer duration.");
                return;
            }

            // Stop any running timer
            stopTimer();

            // Calculate total seconds
            timeLeft = (hours * 3600) + (minutes * 60) + seconds;
            initialTimeLeft = timeLeft; // Store for reset
            updateTimerDisplay();

            // Clear input fields
            if (hoursInput) hoursInput.value = "";
            if (minutesInput) minutesInput.value = "";
            if (secondsInput) secondsInput.value = "";
        }

        function startTimer() {
            console.log("startTimer called");
            if (timeLeft <= 0) {
                alert("Please set a timer duration first.");
                return;
            }
            if (!timer) {
                timer = setInterval(() => {
                    if (timeLeft <= 0) {
                        clearInterval(timer);
                        timer = null;
                        alert("Timer session finished!");
                        resetTimer();
                        return;
                    }
                    timeLeft--;
                    updateTimerDisplay();
                }, 1000);
            }
        }

        function stopTimer() {
            console.log("stopTimer called");
            clearInterval(timer);
            timer = null;
        }

        function resetTimer() {
            console.log("resetTimer called");
            stopTimer();
            timeLeft = initialTimeLeft;
            updateTimerDisplay();
        }

        function updateTimerDisplay() {
            console.log("updateTimerDisplay called");
            const hours = Math.floor(timeLeft / 3600);
            const minutes = Math.floor((timeLeft % 3600) / 60);
            const seconds = timeLeft % 60;
            const timerDisplay = document.getElementById("timerDisplay");
            if (timerDisplay) {
                timerDisplay.textContent =
                    `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
            }
        }

        // Initialize
        console.log("Initializing application");
        loadData();
        displayQuote();
        displaySubjects();
        displayTasks();
        generateCalendar();
        updateOverallScore();
        updateTimerDisplay();

function initializeApp() {
    loadData();
    displayQuote();
    displaySubjects();
    displayTasks();
    generateCalendar();
    updateOverallScore();
    updateTimerDisplay();
}
