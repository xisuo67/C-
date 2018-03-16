ImportFunction = (function () {
    initGrid = function () {
        GridList = $("#ListTable").jqGrid({
            url: "../Import/Query",
            datatype: "json",
            colNames: ['学生编号', '学生姓名', '年龄', '性别'],
            colModel: [
                { name: 'SNO', index: 'SNO', width: 55 ,align: 'center'},
                { name: 'SNAME', index: 'SNAME', width: 90 ,align: 'center'},
                { name: 'AGE', index: 'AGE', width: 100, align: 'center'},
                { name: 'SEX', index: 'SEX', width: 80, align: "center" }
            ],
            sortorder: "desc",
        });
    };
    return {
        init: function () {
            initGrid();
        }
    }
})();