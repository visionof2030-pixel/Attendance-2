
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>أداة تحضير الطلاب للمعلمين</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        :root {
            --primary-color: #4a6fa5;
            --secondary-color: #6b8cbc;
            --accent-color: #f9a826;
            --light-color: #f8f9fa;
            --dark-color: #343a40;
            --success-color: #28a745;
            --danger-color: #dc3545;
            --border-radius: 8px;
            --box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
            --transition: all 0.3s ease;
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: 'Cairo', sans-serif;
        }

        body {
            background-color: #f0f5ff;
            color: var(--dark-color);
            line-height: 1.6;
            padding: 20px;
            min-height: 100vh;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }

        header {
            text-align: center;
            margin-bottom: 30px;
            padding: 20px;
            background-color: white;
            border-radius: var(--border-radius);
            box-shadow: var(--box-shadow);
        }

        h1 {
            color: var(--primary-color);
            margin-bottom: 10px;
            font-weight: 700;
            font-size: 2.2rem;
        }

        .subtitle {
            color: var(--secondary-color);
            font-size: 1.1rem;
        }

        .main-content {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 25px;
            margin-bottom: 30px;
        }

        @media (max-width: 992px) {
            .main-content {
                grid-template-columns: 1fr;
            }
        }

        .card {
            background-color: white;
            border-radius: var(--border-radius);
            box-shadow: var(--box-shadow);
            padding: 25px;
            transition: var(--transition);
        }

        .card:hover {
            box-shadow: 0 6px 16px rgba(0, 0, 0, 0.12);
        }

        .card-title {
            color: var(--primary-color);
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #eee;
            font-weight: 600;
            font-size: 1.4rem;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .card-title i {
            color: var(--accent-color);
        }

        .form-group {
            margin-bottom: 20px;
        }

        label {
            display: block;
            margin-bottom: 8px;
            color: var(--dark-color);
            font-weight: 500;
        }

        input, select, textarea {
            width: 100%;
            padding: 12px 15px;
            border: 1px solid #ddd;
            border-radius: var(--border-radius);
            font-size: 1rem;
            transition: var(--transition);
        }

        input:focus, select:focus, textarea:focus {
            outline: none;
            border-color: var(--primary-color);
            box-shadow: 0 0 0 3px rgba(74, 111, 165, 0.2);
        }

        .btn {
            padding: 12px 25px;
            border: none;
            border-radius: var(--border-radius);
            font-size: 1rem;
            font-weight: 500;
            cursor: pointer;
            transition: var(--transition);
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
        }

        .btn-primary {
            background-color: var(--primary-color);
            color: white;
        }

        .btn-primary:hover {
            background-color: var(--secondary-color);
        }

        .btn-success {
            background-color: var(--success-color);
            color: white;
        }

        .btn-success:hover {
            background-color: #218838;
        }

        .btn-warning {
            background-color: var(--accent-color);
            color: white;
        }

        .btn-warning:hover {
            background-color: #e69500;
        }

        .btn-danger {
            background-color: var(--danger-color);
            color: white;
        }

        .btn-danger:hover {
            background-color: #c82333;
        }

        .btn-group {
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
            margin-top: 15px;
        }

        .student-list-container {
            overflow-x: auto;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
        }

        th, td {
            padding: 12px 15px;
            text-align: center;
            border-bottom: 1px solid #eee;
        }

        th {
            background-color: #f8f9fa;
            color: var(--primary-color);
            font-weight: 600;
        }

        tr:hover {
            background-color: #f8f9fa;
        }

        .student-count {
            background-color: #e9ecef;
            padding: 8px 15px;
            border-radius: var(--border-radius);
            display: inline-block;
            margin-bottom: 15px;
            font-weight: 600;
            color: var(--dark-color);
        }

        .actions {
            display: flex;
            gap: 8px;
            justify-content: center;
        }

        .action-btn {
            padding: 5px 10px;
            border-radius: 4px;
            border: none;
            cursor: pointer;
            font-size: 0.9rem;
            transition: var(--transition);
        }

        .edit-btn {
            background-color: #e7f1ff;
            color: var(--primary-color);
        }

        .edit-btn:hover {
            background-color: #d0e3ff;
        }

        .delete-btn {
            background-color: #ffeaea;
            color: var(--danger-color);
        }

        .delete-btn:hover {
            background-color: #ffd6d6;
        }

        .empty-state {
            text-align: center;
            padding: 40px 20px;
            color: #6c757d;
        }

        .empty-state i {
            font-size: 3rem;
            margin-bottom: 15px;
            color: #dee2e6;
        }

        .export-options {
            display: flex;
            flex-wrap: wrap;
            gap: 15px;
            margin-top: 25px;
        }

        .export-btn {
            flex: 1;
            min-width: 200px;
        }

        .instructions {
            background-color: #f8f9fa;
            padding: 15px;
            border-radius: var(--border-radius);
            margin-top: 20px;
            border-right: 4px solid var(--accent-color);
        }

        .instructions h3 {
            color: var(--primary-color);
            margin-bottom: 10px;
            font-size: 1.2rem;
        }

        .instructions ul {
            padding-right: 20px;
        }

        .instructions li {
            margin-bottom: 8px;
        }

        .notification {
            position: fixed;
            bottom: 20px;
            left: 20px;
            padding: 15px 20px;
            border-radius: var(--border-radius);
            color: white;
            font-weight: 500;
            z-index: 1000;
            opacity: 0;
            transform: translateY(20px);
            transition: var(--transition);
            max-width: 350px;
        }

        .notification.show {
            opacity: 1;
            transform: translateY(0);
        }

        .notification.success {
            background-color: var(--success-color);
        }

        .notification.error {
            background-color: var(--danger-color);
        }

        .notification.warning {
            background-color: var(--accent-color);
        }

        .footer {
            text-align: center;
            margin-top: 40px;
            padding-top: 20px;
            border-top: 1px solid #ddd;
            color: #6c757d;
            font-size: 0.9rem;
        }

        @media (max-width: 768px) {
            .container {
                padding: 10px;
            }
            
            h1 {
                font-size: 1.8rem;
            }
            
            .card {
                padding: 20px 15px;
            }
            
            .btn-group {
                flex-direction: column;
            }
            
            .btn {
                width: 100%;
            }
            
            .export-btn {
                min-width: 100%;
            }
            
            th, td {
                padding: 10px 8px;
                font-size: 0.9rem;
            }
        }

        .modal {
            display: none;
            position: fixed;
            top: 0;
            right: 0;
            bottom: 0;
            left: 0;
            background-color: rgba(0, 0, 0, 0.5);
            z-index: 1001;
            align-items: center;
            justify-content: center;
        }

        .modal-content {
            background-color: white;
            border-radius: var(--border-radius);
            padding: 30px;
            width: 90%;
            max-width: 500px;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.2);
        }

        .modal-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
        }

        .modal-title {
            color: var(--primary-color);
            font-size: 1.5rem;
            font-weight: 600;
        }

        .close-btn {
            background: none;
            border: none;
            font-size: 1.5rem;
            cursor: pointer;
            color: #6c757d;
        }

        .modal-footer {
            display: flex;
            justify-content: flex-end;
            gap: 10px;
            margin-top: 25px;
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1><i class="fas fa-chalkboard-teacher"></i> أداة تحضير الطلاب للمعلمين</h1>
            <p class="subtitle">أداة متكاملة لإدارة قوائم الطلاب وتوزيعهم على الفصول وتصدير البيانات بتنسيقات مختلفة</p>
        </header>

        <div class="main-content">
            <div class="card">
                <h2 class="card-title"><i class="fas fa-user-tie"></i> بيانات المعلم والفصل</h2>
                <div class="form-group">
                    <label for="teacherName"><i class="fas fa-signature"></i> اسم المعلم</label>
                    <input type="text" id="teacherName" placeholder="أدخل اسمك الكامل">
                </div>
                
                <div class="form-group">
                    <label for="classSelect"><i class="fas fa-door-open"></i> اختر الفصل</label>
                    <select id="classSelect">
                        <option value="الفصل الأول">الفصل الأول</option>
                        <option value="الفصل الثاني">الفصل الثاني</option>
                        <option value="الفصل الثالث">الفصل الثالث</option>
                        <option value="الفصل الرابع">الفصل الرابع</option>
                        <option value="الفصل الخامس">الفصل الخامس</option>
                        <option value="الفصل السادس">الفصل السادس</option>
                    </select>
                </div>
                
                <div class="instructions">
                    <h3><i class="fas fa-info-circle"></i> تعليمات</h3>
                    <ul>
                        <li>أدخل اسمك واختر الفصل قبل إضافة الطلاب</li>
                        <li>يمكنك إضافة طلاب بشكل فردي أو دفعة واحدة</li>
                        <li>لإضافة دفعة طلاب، اكتب أسماءهم كل اسم في سطر جديد</li>
                        <li>يمكنك تعديل أو حذف أي طالب من القائمة</li>
                        <li>احفظ البيانات للرجوع إليها لاحقاً</li>
                    </ul>
                </div>
            </div>

            <div class="card">
                <h2 class="card-title"><i class="fas fa-user-plus"></i> إضافة الطلاب</h2>
                
                <div class="form-group">
                    <label for="singleStudent"><i class="fas fa-user-graduate"></i> إضافة طالب فردي</label>
                    <div class="btn-group">
                        <input type="text" id="singleStudent" placeholder="اسم الطالب">
                        <button class="btn btn-primary" id="addStudentBtn">
                            <i class="fas fa-plus"></i> إضافة
                        </button>
                    </div>
                </div>
                
                <div class="form-group">
                    <label for="batchStudents"><i class="fas fa-users"></i> إضافة دفعة طلاب</label>
                    <textarea id="batchStudents" rows="5" placeholder="ضع كل اسم طالب في سطر جديد
مثال:
أحمد محمد
فاطمة علي
سالم خالد"></textarea>
                    <button class="btn btn-success" id="addBatchBtn">
                        <i class="fas fa-user-friends"></i> إضافة الدفعة
                    </button>
                </div>
                
                <div class="btn-group">
                    <button class="btn btn-warning" id="clearListBtn">
                        <i class="fas fa-trash-alt"></i> مسح القائمة
                    </button>
                    <button class="btn btn-primary" id="saveDataBtn">
                        <i class="fas fa-save"></i> حفظ البيانات
                    </button>
                    <button class="btn btn-success" id="loadDataBtn">
                        <i class="fas fa-folder-open"></i> تحميل البيانات
                    </button>
                </div>
            </div>
        </div>

        <div class="card">
            <h2 class="card-title"><i class="fas fa-list-ol"></i> قائمة الطلاب</h2>
            <div class="student-count" id="studentCount">عدد الطلاب: 0</div>
            
            <div class="student-list-container">
                <table id="studentsTable">
                    <thead>
                        <tr>
                            <th width="10%">الرقم</th>
                            <th width="35%">اسم الطالب</th>
                            <th width="25%">الفصل</th>
                            <th width="20%">تاريخ الإضافة</th>
                            <th width="10%">الإجراءات</th>
                        </tr>
                    </thead>
                    <tbody id="studentsList">
                        <!-- سيتم ملء هذا الجدول ديناميكياً -->
                    </tbody>
                </table>
                
                <div id="emptyState" class="empty-state">
                    <i class="fas fa-user-graduate"></i>
                    <h3>لا يوجد طلاب مضافة بعد</h3>
                    <p>ابدأ بإضافة الطلاب باستخدام النماذج أعلاه</p>
                </div>
            </div>
            
            <div class="export-options">
                <button class="btn btn-warning export-btn" id="exportPDFBtn">
                    <i class="fas fa-file-pdf"></i> تصدير إلى PDF
                </button>
                <button class="btn btn-success export-btn" id="exportExcelBtn">
                    <i class="fas fa-file-excel"></i> تصدير إلى Excel
                </button>
                <button class="btn btn-primary export-btn" id="exportPrintBtn">
                    <i class="fas fa-print"></i> طباعة القائمة
                </button>
            </div>
        </div>
        
        <div class="footer">
            <p>تم تطوير هذه الأداة لتسهيل عملية تحضير الطلاب وإدارة الفصول الدراسية &copy; 2023</p>
        </div>
    </div>

    <!-- نافذة التعديل -->
    <div id="editModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3 class="modal-title"><i class="fas fa-edit"></i> تعديل بيانات الطالب</h3>
                <button class="close-btn" id="closeModalBtn">&times;</button>
            </div>
            <div class="form-group">
                <label for="editStudentName">اسم الطالب</label>
                <input type="text" id="editStudentName">
            </div>
            <div class="form-group">
                <label for="editStudentClass">الفصل</label>
                <select id="editStudentClass">
                    <option value="الفصل الأول">الفصل الأول</option>
                    <option value="الفصل الثاني">الفصل الثاني</option>
                    <option value="الفصل الثالث">الفصل الثالث</option>
                    <option value="الفصل الرابع">الفصل الرابع</option>
                    <option value="الفصل الخامس">الفصل الخامس</option>
                    <option value="الفصل السادس">الفصل السادس</option>
                </select>
            </div>
            <div class="modal-footer">
                <button class="btn btn-danger" id="cancelEditBtn">إلغاء</button>
                <button class="btn btn-success" id="saveEditBtn">حفظ التغييرات</button>
            </div>
        </div>
    </div>

    <!-- إشعارات -->
    <div id="notification" class="notification"></div>

    <script>
        // البيانات والمتغيرات
        let students = [];
        let currentEditIndex = null;
        
        // عناصر DOM
        const teacherNameInput = document.getElementById('teacherName');
        const classSelect = document.getElementById('classSelect');
        const singleStudentInput = document.getElementById('singleStudent');
        const addStudentBtn = document.getElementById('addStudentBtn');
        const batchStudentsTextarea = document.getElementById('batchStudents');
        const addBatchBtn = document.getElementById('addBatchBtn');
        const clearListBtn = document.getElementById('clearListBtn');
        const saveDataBtn = document.getElementById('saveDataBtn');
        const loadDataBtn = document.getElementById('loadDataBtn');
        const studentsList = document.getElementById('studentsList');
        const emptyState = document.getElementById('emptyState');
        const studentCount = document.getElementById('studentCount');
        const exportPDFBtn = document.getElementById('exportPDFBtn');
        const exportExcelBtn = document.getElementById('exportExcelBtn');
        const exportPrintBtn = document.getElementById('exportPrintBtn');
        const editModal = document.getElementById('editModal');
        const closeModalBtn = document.getElementById('closeModalBtn');
        const cancelEditBtn = document.getElementById('cancelEditBtn');
        const editStudentName = document.getElementById('editStudentName');
        const editStudentClass = document.getElementById('editStudentClass');
        const saveEditBtn = document.getElementById('saveEditBtn');
        const notification = document.getElementById('notification');
        
        // تحميل البيانات المحفوظة عند بدء التشغيل
        document.addEventListener('DOMContentLoaded', () => {
            loadStudentsFromStorage();
            updateStudentCount();
            renderStudentsList();
            
            // إذا كان هناك بيانات محفوظة للمعلم، قم بتحميلها
            const savedTeacher = localStorage.getItem('teacherData');
            if (savedTeacher) {
                const teacherData = JSON.parse(savedTeacher);
                teacherNameInput.value = teacherData.name || '';
                classSelect.value = teacherData.class || 'الفصل الأول';
            }
        });
        
        // إضافة طالب فردي
        addStudentBtn.addEventListener('click', () => {
            const studentName = singleStudentInput.value.trim();
            const className = classSelect.value;
            const teacherName = teacherNameInput.value.trim();
            
            if (!studentName) {
                showNotification('يرجى إدخال اسم الطالب', 'error');
                return;
            }
            
            if (!teacherName) {
                showNotification('يرجى إدخال اسم المعلم أولاً', 'warning');
                return;
            }
            
            addStudent(studentName, className);
            singleStudentInput.value = '';
            showNotification('تم إضافة الطالب بنجاح', 'success');
        });
        
        // إضافة طالب عند الضغط على Enter في حقل الإدخال الفردي
        singleStudentInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                addStudentBtn.click();
            }
        });
        
        // إضافة دفعة طلاب
        addBatchBtn.addEventListener('click', () => {
            const batchText = batchStudentsTextarea.value.trim();
            const className = classSelect.value;
            const teacherName = teacherNameInput.value.trim();
            
            if (!batchText) {
                showNotification('يرجى إدخال قائمة الطلاب', 'error');
                return;
            }
            
            if (!teacherName) {
                showNotification('يرجى إدخال اسم المعلم أولاً', 'warning');
                return;
            }
            
            const studentNames = batchText.split('\n')
                .map(name => name.trim())
                .filter(name => name !== '');
            
            if (studentNames.length === 0) {
                showNotification('لا توجد أسماء طلاب صالحة للإضافة', 'warning');
                return;
            }
            
            studentNames.forEach(name => {
                addStudent(name, className);
            });
            
            batchStudentsTextarea.value = '';
            showNotification(`تم إضافة ${studentNames.length} طالب بنجاح`, 'success');
        });
        
        // مسح القائمة
        clearListBtn.addEventListener('click', () => {
            if (students.length === 0) {
                showNotification('لا يوجد طلاب لحذفهم', 'warning');
                return;
            }
            
            if (confirm('هل أنت متأكد من مسح جميع الطلاب؟ لا يمكن التراجع عن هذا الإجراء.')) {
                students = [];
                updateStudentCount();
                renderStudentsList();
                saveStudentsToStorage();
                showNotification('تم مسح جميع الطلاب بنجاح', 'success');
            }
        });
        
        // حفظ البيانات
        saveDataBtn.addEventListener('click', () => {
            const teacherName = teacherNameInput.value.trim();
            
            if (!teacherName) {
                showNotification('يرجى إدخال اسم المعلم أولاً', 'warning');
                return;
            }
            
            // حفظ بيانات المعلم
            const teacherData = {
                name: teacherName,
                class: classSelect.value
            };
            
            localStorage.setItem('teacherData', JSON.stringify(teacherData));
            
            // حفظ بيانات الطلاب
            saveStudentsToStorage();
            showNotification('تم حفظ البيانات بنجاح', 'success');
        });
        
        // تحميل البيانات
        loadDataBtn.addEventListener('click', () => {
            loadStudentsFromStorage();
            updateStudentCount();
            renderStudentsList();
            showNotification('تم تحميل البيانات بنجاح', 'success');
        });
        
        // تصدير إلى PDF
        exportPDFBtn.addEventListener('click', () => {
            if (students.length === 0) {
                showNotification('لا يوجد بيانات للتصدير', 'warning');
                return;
            }
            
            const teacherName = teacherNameInput.value.trim() || 'غير محدد';
            const className = classSelect.value;
            
            // إنشاء مستند PDF
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF();
            
            // إضافة عنوان التقرير
            doc.setFontSize(20);
            doc.setTextColor(74, 111, 165);
            doc.text('قائمة الطلاب', 105, 15, null, null, 'center');
            
            doc.setFontSize(12);
            doc.setTextColor(0, 0, 0);
            doc.text(`المعلم: ${teacherName}`, 15, 25);
            doc.text(`الفصل: ${className}`, 15, 32);
            doc.text(`تاريخ التصدير: ${new Date().toLocaleDateString('ar-SA')}`, 15, 39);
            
            // إعداد جدول البيانات
            const headers = [['الرقم', 'اسم الطالب', 'الفصل', 'تاريخ الإضافة']];
            const data = students.map((student, index) => [
                index + 1,
                student.name,
                student.class,
                student.date
            ]);
            
            // رسم الجدول
            doc.autoTable({
                head: headers,
                body: data,
                startY: 45,
                theme: 'grid',
                headStyles: { fillColor: [74, 111, 165], textColor: [255, 255, 255] },
                styles: { font: 'Cairo', fontSize: 10, halign: 'center' },
                margin: { right: 15, left: 15 }
            });
            
            // حفظ الملف
            doc.save(`طلاب_${className}_${new Date().toISOString().slice(0,10)}.pdf`);
            showNotification('تم تصدير الملف PDF بنجاح', 'success');
        });
        
        // تصدير إلى Excel
        exportExcelBtn.addEventListener('click', () => {
            if (students.length === 0) {
                showNotification('لا يوجد بيانات للتصدير', 'warning');
                return;
            }
            
            const teacherName = teacherNameInput.value.trim() || 'غير محدد';
            const className = classSelect.value;
            
            // إعداد البيانات
            const header = ['الرقم', 'اسم الطالب', 'الفصل', 'تاريخ الإضافة'];
            const data = students.map((student, index) => [
                index + 1,
                student.name,
                student.class,
                student.date
            ]);
            
            // إضافة معلومات إضافية
            const infoData = [
                ['قائمة الطلاب'],
                [`المعلم: ${teacherName}`],
                [`الفصل: ${className}`],
                [`تاريخ التصدير: ${new Date().toLocaleDateString('ar-SA')}`],
                [],
                ...data
            ];
            
            // إنشاء مصنف Excel
            const ws = XLSX.utils.aoa_to_sheet([header, ...data]);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'قائمة الطلاب');
            
            // حفظ الملف
            XLSX.writeFile(wb, `طلاب_${className}_${new Date().toISOString().slice(0,10)}.xlsx`);
            showNotification('تم تصدير الملف Excel بنجاح', 'success');
        });
        
        // طباعة القائمة
        exportPrintBtn.addEventListener('click', () => {
            if (students.length === 0) {
                showNotification('لا يوجد بيانات للطباعة', 'warning');
                return;
            }
            
            const printWindow = window.open('', '_blank');
            const teacherName = teacherNameInput.value.trim() || 'غير محدد';
            const className = classSelect.value;
            
            let printContent = `
                <!DOCTYPE html>
                <html dir="rtl">
                <head>
                    <meta charset="UTF-8">
                    <title>قائمة الطلاب - ${className}</title>
                    <style>
                        body { font-family: 'Cairo', sans-serif; padding: 20px; }
                        h1 { color: #4a6fa5; text-align: center; }
                        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
                        th, td { padding: 10px; border: 1px solid #ddd; text-align: center; }
                        th { background-color: #4a6fa5; color: white; }
                        .info { margin-bottom: 20px; }
                        @media print {
                            .no-print { display: none; }
                        }
                    </style>
                </head>
                <body>
                    <h1>قائمة الطلاب</h1>
                    <div class="info">
                        <p><strong>المعلم:</strong> ${teacherName}</p>
                        <p><strong>الفصل:</strong> ${className}</p>
                        <p><strong>تاريخ الطباعة:</strong> ${new Date().toLocaleDateString('ar-SA')}</p>
                    </div>
                    <table>
                        <thead>
                            <tr>
                                <th>الرقم</th>
                                <th>اسم الطالب</th>
                                <th>الفصل</th>
                                <th>تاريخ الإضافة</th>
                            </tr>
                        </thead>
                        <tbody>`;
            
            students.forEach((student, index) => {
                printContent += `
                    <tr>
                        <td>${index + 1}</td>
                        <td>${student.name}</td>
                        <td>${student.class}</td>
                        <td>${student.date}</td>
                    </tr>`;
            });
            
            printContent += `
                        </tbody>
                    </table>
                    <button class="no-print" onclick="window.print()" style="margin-top: 20px; padding: 10px 20px; background: #4a6fa5; color: white; border: none; border-radius: 5px; cursor: pointer;">
                        طباعة
                    </button>
                </body>
                </html>`;
            
            printWindow.document.write(printContent);
            printWindow.document.close();
            showNotification('تم فتح نافذة الطباعة', 'success');
        });
        
        // إغلاق نافذة التعديل
        closeModalBtn.addEventListener('click', () => {
            editModal.style.display = 'none';
        });
        
        cancelEditBtn.addEventListener('click', () => {
            editModal.style.display = 'none';
        });
        
        // حفظ التعديلات
        saveEditBtn.addEventListener('click', () => {
            if (currentEditIndex === null) return;
            
            const newName = editStudentName.value.trim();
            const newClass = editStudentClass.value;
            
            if (!newName) {
                showNotification('يرجى إدخال اسم الطالب', 'error');
                return;
            }
            
            students[currentEditIndex].name = newName;
            students[currentEditIndex].class = newClass;
            
            updateStudentCount();
            renderStudentsList();
            saveStudentsToStorage();
            editModal.style.display = 'none';
            
            showNotification('تم تعديل بيانات الطالب بنجاح', 'success');
        });
        
        // وظائف مساعدة
        function addStudent(name, className) {
            const newStudent = {
                id: Date.now(),
                name: name,
                class: className,
                date: new Date().toLocaleDateString('ar-SA')
            };
            
            students.push(newStudent);
            updateStudentCount();
            renderStudentsList();
        }
        
        function updateStudentCount() {
            studentCount.textContent = `عدد الطلاب: ${students.length}`;
            
            // إظهار أو إخفاء حالة القائمة الفارغة
            if (students.length === 0) {
                emptyState.style.display = 'block';
                studentsList.style.display = 'none';
            } else {
                emptyState.style.display = 'none';
                studentsList.style.display = 'table-row-group';
            }
        }
        
        function renderStudentsList() {
            studentsList.innerHTML = '';
            
            students.forEach((student, index) => {
                const row = document.createElement('tr');
                
                row.innerHTML = `
                    <td>${index + 1}</td>
                    <td>${student.name}</td>
                    <td>${student.class}</td>
                    <td>${student.date}</td>
                    <td>
                        <div class="actions">
                            <button class="action-btn edit-btn" data-index="${index}">
                                <i class="fas fa-edit"></i>
                            </button>
                            <button class="action-btn delete-btn" data-index="${index}">
                                <i class="fas fa-trash"></i>
                            </button>
                        </div>
                    </td>
                `;
                
                studentsList.appendChild(row);
            });
            
            // إضافة مستمعي الأحداث لأزرار التعديل والحذف
            document.querySelectorAll('.edit-btn').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    const index = parseInt(e.currentTarget.getAttribute('data-index'));
                    openEditModal(index);
                });
            });
            
            document.querySelectorAll('.delete-btn').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    const index = parseInt(e.currentTarget.getAttribute('data-index'));
                    deleteStudent(index);
                });
            });
        }
        
        function openEditModal(index) {
            currentEditIndex = index;
            const student = students[index];
            
            editStudentName.value = student.name;
            editStudentClass.value = student.class;
            editModal.style.display = 'flex';
        }
        
        function deleteStudent(index) {
            if (confirm('هل أنت متأكد من حذف هذا الطالب؟')) {
                students.splice(index, 1);
                updateStudentCount();
                renderStudentsList();
                saveStudentsToStorage();
                showNotification('تم حذف الطالب بنجاح', 'success');
            }
        }
        
        function saveStudentsToStorage() {
            localStorage.setItem('studentsData', JSON.stringify(students));
        }
        
        function loadStudentsFromStorage() {
            const savedStudents = localStorage.getItem('studentsData');
            if (savedStudents) {
                students = JSON.parse(savedStudents);
            }
        }
        
        function showNotification(message, type) {
            notification.textContent = message;
            notification.className = `notification ${type} show`;
            
            setTimeout(() => {
                notification.classList.remove('show');
            }, 3000);
        }
        
        // إضافة تأثيرات تفاعلية
        document.querySelectorAll('.card').forEach(card => {
            card.addEventListener('mouseenter', () => {
                card.style.transform = 'translateY(-5px)';
            });
            
            card.addEventListener('mouseleave', () => {
                card.style.transform = 'translateY(0)';
            });
        });
        
        // إضافة تأثيرات للأزرار
        document.querySelectorAll('.btn').forEach(btn => {
            btn.addEventListener('mousedown', () => {
                btn.style.transform = 'scale(0.98)';
            });
            
            btn.addEventListener('mouseup', () => {
                btn.style.transform = 'scale(1)';
            });
            
            btn.addEventListener('mouseleave', () => {
                btn.style.transform = 'scale(1)';
            });
        });
    </script>
</body>
</html>
