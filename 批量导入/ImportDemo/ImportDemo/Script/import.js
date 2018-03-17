ImportFunction = (function () {
    bindEvent = function () {
        //批量导出
        $("#btn_Export").click( function () {
            alert(2)
            Export("学生信息");
        });
        //批量导入
        $("#btn_Import_Template").on("click", function () {
            alert(1)
        });
    };
    ExportMore = function (param) {
        debugger
        if (param) {
            //var columnInfos = this.getColumnInfo(param.columnFilter);
            var options = {
                //columnInfos: columnInfos,
                fileName: param.FileName,
                type: type,
                FileFormat: fileFormat,
                url: param.url,
                FixColumns: param.FixColumns,
                GroupHeader: param.GroupHeader,
                Tag: param.Tag,
                ColAsSerialize: param.ColAsSerialize
            };
            //默认不分页导出数据
            if (!$.isArray(param.Data)) {
                options.condition = this.getLoadParams();
                options.api = settings.url;
            } else {
                options.data = param.Data;
            }
            var guid = UUID();//生成Guid,服务端以此跟踪进度
            ExportXLS(options, guid);
        }
    };
    UUID = function () {
        var s = [], itoh = '0123456789ABCDEF'.split('');
        for (var i = 0; i < 36; i++) s[i] = Math.floor(Math.random() * 0x10);
        s[14] = 4;
        s[19] = (s[19] & 0x3) | 0x8;
        for (var i = 0; i < 36; i++) s[i] = itoh[s[i]];
        s[8] = s[13] = s[18] = s[23] = '-';
        return s.join('');
    };
    Export = function (fileName, data, url) {
        this.ExportMore({
            FileName: fileName,
            Data: data,
            url: url
        });
    };
    ExportXLS= function (options, guid) {
        if (options.columnInfos && options.columnInfos.length > 0) {
            if ($("#excelForm").length == 0) {
                $('<form id="excelForm"  method="post" target="excelIFrame"><input type="hidden" name="excelParam" id="excelData" /></form><iframe id="excelIFrame" name="excelIFrame" style="display:none;"></iframe>').appendTo("body");
            }
            if (options.FileFormat == "undefined") {
                options.FileFormat = "0";
            }
            if (!options.url) {
                options.url = "../Excel/GridExport";
            }
            if (options.ColAsSerialize == undefined || options.ColAsSerialize == null) {
                options.ColAsSerialize = true;
            }
            var param = {
                "FileName": options.fileName,
                "ColumnInfoList": options.columnInfos,
                "Type": options.type,
                "FileFormat": options.FileFormat,
                "Remark": options.Remark,
                "Tag": options.Tag,
                "FixColumns": options.FixColumns,
                "GroupHeader": options.GroupHeader,
                "ColAsSerialize": options.ColAsSerialize
            };
            //var strFunctionCode = "";
            //if ($.isFunction(Permission.GetFuncitonCode)) {
            //    param.FunctionCode = Permission.GetFuncitonCode();
            //}
            //默认不分页导出数据
            if (!$.isArray(options.data)) {
                //查询条件
                var queryParam = citms.clone(options.condition);
                queryParam.page = 1;

                param.Condition = queryParam;
                param.RootPath = document.location.origin;
                param.Api = document.location.origin + options.api;
                param.IsExportSelectData = false;
            } else {
                param.Data = options.data;
                param.IsExportSelectData = true;
            }

            if (guid) {
                param.Guid = guid;
            }

            $("#excelData").val(JSON.stringify(param));
            var excelForm = document.getElementById("excelForm");

            excelForm.action = options.url;
            excelForm.submit();
        }
    };
    initData = function () {
        $.ajax({
            type: "post",
            url: "../Import/Query",
            dataType: "json",
            success: function (data) {
                if (!data.Result) {
                    layer.msg(data.Message, { icon: 2 });
                } else {
                    initTable(data.Result);
                }
            }
        });
    };
    initTable = function (studentsInfo) {
        var vm = new Vue({
            el: '#app',
            data: {
                Students: studentsInfo
            },
        });
    };
    return {
        init: function () {
            bindEvent();
            initData();
        }
    }
})();