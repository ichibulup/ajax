
{/* <reference path="jquery-3.7.1.intellisense.js" /> */}

$(document).ready(() => {
    loadData();

    if ($.fn.DataTable.isDataTable('#employeeTable')) {
        $('#employeeTable').DataTable().destroy();
        console.log("Table is exist")
    }

    $('#addBtn').click(() => {
        clearModal();
        console.log("Open modal");
        $('#employeeModal').modal('show');
    });

    $('#saveBtn').click(() => {
        let id = $('#employeeId').val();
        let employee = {
            EmployeeID: id,
            Name: $('#employeeName').val(),
            Age: $('#employeeAge').val(),
            State: $('#employeeState').val(),
            Country: $('#employeeCountry').val()
        };

        if (id) {
            updateEmployee(employee);
        } else {
            addEmployee(employee);
        }
    });

    $('#exportExcelBtn').click(() => {
        window.location.href = '/Home/ExportToExcel';
    });

    $('#exportWordBtn').click(() => {
        window.location.href = '/Home/ExportToWord';
    });
});

let loadData = () => {
    console.log("Initializing DataTable");
    // $('#employeeTable').empty();
    $('#employeeTable').DataTable({
        "processing": true,
        "serverSide": true,
        //"filter": true,
        "paging": true,
        "searching": false,
        // "orderMulti": false,
        "pageLength": 10,
        ajax: {
            url: '/Home/List',
            dataType: "json",
            // contentType: "application/json;charset=utf-8",
            type: 'POST'
        },
        columns: [
            { data: 'EmployeeID' },
            { data: 'Name' },
            { data: 'Age' },
            { data: 'State' },
            { data: 'Country' },
            {
                data: 'EmployeeID',
                render: (data) => {
                    return `
                        <button class="btn btn-warning" onclick="editEmployee(${data})">Edit</button>
                        <button class="btn btn-danger" onclick="deleteEmployee(${data})">Delete</button>
                    `;
                }
            }
        ]
    });
    console.log("Done");
}

let addEmployee = (employee) => {
    $.post('/Home/Add', employee, () => {
        $('#employeeModal').modal('hide');
        $('#employeeTable').DataTable().ajax.reload();
    });
}

let updateEmployee = (employee) => {
    $.post('/Home/Update', employee, () => {
        $('#employeeModal').modal('hide');
        $('#employeeTable').DataTable().ajax.reload();
    });
}

let deleteEmployee = (id) => {
    $.post('/Home/Delete', { ID: id }, () => {
        $('#employeeTable').DataTable().ajax.reload();
    });
}

let editEmployee = (id) => {
    $.get('/Home/GetbyID', { ID: id }, (data) => {
        $('#employeeId').val(data.EmployeeID);
        $('#employeeName').val(data.Name);
        $('#employeeAge').val(data.Age);
        $('#employeeState').val(data.State);
        $('#employeeCountry').val(data.Country);
        $('#employeeModal').modal('show');
    });
}

let clearModal = () => {
    $('#employeeId').val('');
    $('#employeeName').val('');
    $('#employeeAge').val('');
    $('#employeeState').val('');
    $('#employeeCountry').val('');
}
