document.getElementById('submit').addEventListener('click', function() {
    readExcelFile(function(data) {
        // console.log(data);
        create_student_folder(data);
    });
});

function readExcelFile(callback) {
    const fileInput = document.getElementById('excel-file');
    const file = fileInput.files[0];
    if (!file) {
        alert('لطفا فایل اکسل را انتخاب کنید')
        return;
    }
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        const keys = jsonData[0];
        const result = jsonData.slice(1).map(row => {
            return keys.reduce((obj, key, index) => {
                obj[key] = row[index] || null;
                return obj;
            }, {});
        });
        callback(result);
    };
    reader.readAsArrayBuffer(file);
}


function create_student_folder(data) {
    const selected_grade = document.getElementById("select-grade").value;
    let $iframe = $('#folders-iframe'); // Use jQuery to select the iframe
    let iframeDoc = $iframe[0].contentDocument || $iframe[0].contentWindow.document; // Access the iframe document
    let $foldersBody = $(iframeDoc).find('#body'); // Use jQuery to find the 'body' element in the iframe
    $foldersBody.empty();
    Object.entries(data).forEach(([, student]) => {
        if(selected_grade === 'همه' || selected_grade === student['پایه']) {
            create_cover_page(student);
            create_parents_evaluation_page(student);
            create_self_evaluation_page(student);
            create_absences_page(student);
            create_disorders_page(student);
            create_manager_evaluation_page(student);
            create_teacher_evaluation_page(student);

            $(iframeDoc).find('#print-button').click();
        }
    });
}

function create_cover_page(student) {
    let cover_template = get_page_template('cover-page-template');
    const page_content = write_vars(cover_template, [
        'province-name',
        'city-name',
        'student-name',
        'grade',
        'school-name',
        'academic-year'
    ], [
        $('#province-name').val(),
        $('#city-name').val(),
        student['نام']+ ' ' +student['نام خانوادگي'],
        student['پایه'],
        $('#school-name').val(),
        $('#select-academic-year').val()
    ]);
    return add_page_to_folders(page_content);
}


function create_parents_evaluation_page(student) {
    ['پائیز', 'زمستان', 'بهار'].forEach((season) => {
        let page_template = get_page_template('parents-evaluation-page-template');
        const page_content = write_vars(page_template, [
            'student-name',
            'grade',
            'school-name',
            'season'
        ], [
            student['نام']+ ' ' +student['نام خانوادگي'],
            student['پایه'],
            $('#school-name').val(),
            season
        ]);
        add_page_to_folders(page_content);
    });
}

function create_self_evaluation_page(student) {
    [1, 2].forEach((page_number) => {
        let page_template = get_page_template('self-evaluation-page-template');
        const page_content = write_vars(page_template, [
            'student-name',
            'grade',
            'school-name',
            'self-evaluation-page-number',
            'first-col-months-name',
            'second-col-months-name'
        ], [
            student['نام']+ ' ' +student['نام خانوادگي'],
            student['پایه'],
            $('#school-name').val(),
            page_number,
            (page_number === 1) ? 'مهر و آبان' : 'بهمن و اسفند',
            (page_number === 1) ? 'آذر و دی' : 'فروردین و اردیبهشت',
        ]);
        add_page_to_folders(page_content);
    });
}

function create_absences_page(student) {
    let page_template = get_page_template('students-absences-page-template');
    const page_content = write_vars(page_template, [
        'student-name',
        'grade',
        'school-name',
    ], [
        student['نام']+ ' ' +student['نام خانوادگي'],
        student['پایه'],
        $('#school-name').val(),
    ]);
    return add_page_to_folders(page_content);
}

function create_disorders_page(student) {
    ['تحصیلی', 'انضباطی'].forEach((disorder_type) => {
        let page_template = get_page_template('disorders-page-template');
        const page_content = write_vars(page_template, [
            'disorder-type',
            'student-name',
            'grade',
            'school-name',
        ], [
            disorder_type,
            student['نام']+ ' ' +student['نام خانوادگي'],
            student['پایه'],
            $('#school-name').val(),
        ]);
        add_page_to_folders(page_content);
    });
}

function create_manager_evaluation_page(student) {
    let page_template = get_page_template('manager-evaluation-page-template');
    const page_content = write_vars(page_template, [
        'academic-year',
        'student-name',
        'grade',
        'school-name',
    ], [
        $('#select-academic-year').val(),
        student['نام']+ ' ' +student['نام خانوادگي'],
        student['پایه'],
        $('#school-name').val(),
    ]);
    add_page_to_folders(page_content);
}

function create_teacher_evaluation_page(student) {
    const grade_lessons = {
        'اول' : ['قرآن', 'فارسی', 'ریاضی', 'علوم', 'هنر', 'ورزش', 'شایستگی'],
        'دوم' : ['شایستگی', 'هدیه‌آسمانی', 'ورزش', 'هنر', 'علوم', 'ریاضی', 'فارسی', 'قرآن'],
        'سوم' : ['شایستگی', 'هدیه‌آسمانی', 'ورزش', 'هنر', 'علوم', 'ریاضی', 'فارسی', 'قرآن', 'مطالعات‌اجتماعی'],
        'چهارم' : ['شایستگی', 'هدیه‌آسمانی', 'ورزش', 'هنر', 'علوم', 'ریاضی', 'فارسی', 'قرآن', 'مطالعات‌اجتماعی'],
        'پنجم' : ['شایستگی', 'هدیه‌آسمانی', 'ورزش', 'هنر', 'علوم', 'ریاضی', 'فارسی', 'قرآن', 'مطالعات‌اجتماعی'],
        'ششم' : ['شایستگی', 'هدیه‌آسمانی', 'ورزش', 'هنر', 'علوم', 'ریاضی', 'فارسی', 'قرآن', 'مطالعات‌اجتماعی', 'تفکر و پژوهش', 'کار و فناوری']
    };

    ['مهر', 'آبان', 'آذر', 'بهمن', 'اسفند', 'فروردین', 'اردیبهشت'].forEach((month) => {
        let page_template = get_page_template('teacher-evaluation-page-template');
        let lesson_rows = '';
        let i = 0;
        grade_lessons[student['پایه']].forEach((lesson_name) => {
            i++;
            lesson_rows += `<tr>
                <td>${i}</td>
                <td>${lesson_name}</td>
                <td></td>
                <td></td>
             </tr>`;
        });
        const page_content = write_vars(page_template, [
            'month',
            'student-name',
            'grade',
            'school-name',
            'lessons',
        ], [
            month,
            student['نام']+ ' ' +student['نام خانوادگي'],
            student['پایه'],
            $('#school-name').val(),
            lesson_rows
        ]);
        add_page_to_folders(page_content);
    });

}

function get_page_template(page_name) {
    let iframe = $('#template-iframe').contents();
    return iframe.find('#' + page_name).clone();
}

function add_page_to_folders(page) {
    return $('#folders-iframe').contents().find("#body").append(page);
}

function write_vars(page_content, var_name, var_value) {
    let i = 0;
    var_name.forEach((variable) => {
        $(page_content).find('.' + variable).each(function() {
            $(this).html(var_value[i++]);
        });
    });
    return page_content;
}