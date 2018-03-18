ImportFunction = (function () {
    bindEvent = function () {
        //批量导出
        $("#btn_Export").click( function () {
            Export("学生信息");
        });
        //批量导入
        $("#btn_Import_Template").on("click", function () {
            ImportExcelTemplate({
                type: "StudentsInfo", after: function () {
                    initData();
                }
            });
        });
    };
    loadFile=function (name) {
        var js = /js$/i.test(name);
        var bInclude = false;
        var tag = js ? 'script' : 'link';
        var attr = js ? 'src' : 'href';
        var es = document.getElementsByTagName(tag);
        for (var i = 0; i < es.length; i++) {
            if (es[i][attr].indexOf(name) != -1) {
                bInclude = true;
                break;
            }
        }
        if (!bInclude) {
            $(js ? 'body' : 'head').append('<{0} {3} {1}="{2}"></{0}>'.format(tag, attr, name, js ? 'type="text/javascript"' : 'rel="stylesheet"'));
        }
    };
    /*
  * 功能：    根据业务类型下载导入数据得模版文件
  * 参数：    options：
              {
                  type:业务类型, 取值参照 Utility.ExcelImportType 枚举
                  Ext:可导入文件类型,
                  ReturnDetailData:是否返回详细数据
                  after:function(){}//回调函数
              }
  * 返回值：  无
  * 创建人：  杜冬军
  */
    ImportExcelTemplate = function (options) {
        if ($.isPlainObject(options)) {
            var defaults = {
                ReturnDetailData: 0
            };
            var param = $.extend({}, defaults, options);
            if (param.type != "undefined") {
                //加载样式和js文件
                loadFile("../CSS/webuploader/webuploader.css");
                loadFile("../Script/webuploader/webuploader.min.js");
                if (!WebUploader.Uploader.support()) {
                    var error = "上传控件不支持您的浏览器！请尝试升级flash版本或者使用Chrome引擎的浏览器";
                    if (window.console) {
                        window.console.log(error);
                    }
                    return;
                }

                var id = "ImportExcelTemplate_Single";
                var pickerId = "picker_{0}".format(param.type);
                var modal = $("#" + id), loadIndex;
                loadIndex = layer.load(0);
                $(modal).remove();
                var html =
                    '<div class="modal" id="{0}">'.format(id) +
                    '<div class="modal-dialog">' +
                    '<div class="modal-content">' +
                    '<div class="modal-header">' +
                    '<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>' +
                    '<h4 class="modal-title">Excel导入</h4>' +
                    '</div>' +
                    '<div class="modal-body">' +
                    '<div id="uploader" class="wu-example">' +
                    '<p style="font-weight:bold;">导入说明:</p><p class="pt5">导入文件为EXCEL格式，请先下载模板进行必要信息填写，模板下载<a href="javascript:;" onclick="$.DownloadExcelTemplate(\'{0}\')">请点击这里</a>！</p>'.format(param.type) +
                    '<div id="thelist" class="uploader-list"></div>' +
                    '<div class="uploader-wrap clearfix pb20">' +
                    '<input type="text" readonly class="form-control input-sm mr5 upload-file-name" style="width:300px;" />' +
                    '<div id="{0}">选择文件</div>'.format(pickerId) +
                    '<button id="ctlBtn" class="btn btn-white btn-sm btn-start-uploader ml5" style="display:none;">开始上传</button>' +
                    '</div>'
                '</div>' +
                    '</div></div></div></div>';
                $(html).appendTo("body");
                modal = $("#" + id);
                setTimeout(function () {
                    modal.modal('show');
                    var postData = { type: param.type, FunctionCode: param.FunctionCode, ReturnDetailData: param.ReturnDetailData };
                    var uploader = WebUploader.create({
                        swf: '../Scripts/plugins/webuploader/Uploader.swf',
                        server: '../Import/ImportTemplate?' + $.param(postData),
                        pick: '#' + pickerId,
                        accept: {
                            title: 'excel',
                            extensions: 'xls',
                            mimeTypes: 'application/msexcel'
                        },
                        resize: false,
                        fileSingleSizeLimit: 10 * 1024 * 1024,//10M
                        duplicate: true
                    });
                    $("#ctlBtn").on('click', function () {
                        uploader.upload();
                    });

                    // 当有文件被添加进队列的时候
                    uploader.on('fileQueued', function (file) {
                        $("#thelist").html('<div id="' + file.id + '" class="item">' +
                            '<div class="state"></div>' +
                            '</div>');
                        $(".upload-file-name").val(file.name);
                        $(".btn-start-uploader").show();
                    });

                    // 文件上传过程中创建进度条实时显示。
                    uploader.on('uploadProgress', function (file, percentage) {
                        var $li = $('#' + file.id),
                            $percent = $li.find('.progress .progress-bar');

                        // 避免重复创建
                        if (!$percent.length) {
                            $percent = $('<div class="progress progress-striped active">' +
                                '<div class="progress-bar" role="progressbar" style="width: 0%">' +
                                '</div>' +
                                '</div>').appendTo($li).find('.progress-bar');
                        }

                        $li.find('.state').text('上传中');

                        $percent.css('width', percentage * 100 + '%');
                        $(".upload-file-name").val("");
                        $(".btn-start-uploader").hide();
                    });

                    uploader.on('uploadSuccess', function (file, response) {
                        if (response.IsSuccess) {
                            $('#' + file.id).find('.state').html('<span class="label label-success">' + response.Message + '</span>');
                            if ($.isFunction(param.after)) {
                                param.after(response, modal);
                            }
                        } else {
                            if (response.Message.indexOf("http://") >= 0) {
                                $('#' + file.id).find('.state').html("上传的数据中存在错误数据，请点击<a class='red' href='{0}' target='_blank'>下载错误数据</a>！".format(response.Message));
                            } else {
                                $('#' + file.id).find('.state').html('<span class="label label-danger" title="' + response.Message + '">' + response.Message + '</span>');
                            }
                        }


                    });

                    uploader.on('uploadError', function (file, response) {
                        console.log(response);
                        $('#' + file.id).find('.state').text('上传出错');
                    });

                    uploader.on('uploadComplete', function (file) {
                        $('#' + file.id).find('.progress').fadeOut(200);
                    });
                    layer.close(loadIndex);
                }, 300);



            }
        }
    };
    getColumnInfo = function () {
        var tableInfo = $('table thead tr th');
        var columnInfos = new Array();
        var itemColumn = null;
        $.each(tableInfo, function (index, item) {
            itemColumn = {
                "Align": item.align ? item.align : "left",
                "Header": item.title,
                "Field": item.dataset.field,
            };
            columnInfos.push(itemColumn);
        });
        return columnInfos;
    };
    ExportMore = function (param) {
        if (param) {
            var columnInfos = this.getColumnInfo();
            var fileFormat = 0;
            if (param.FileFormat != null) {
                fileFormat = param.FileFormat;
            }
            var options = {
                columnInfos: columnInfos,
                fileName: param.FileName,
                type:"post",
                FileFormat: fileFormat,
                url: param.url,
                FixColumns: param.FixColumns,
                GroupHeader: param.GroupHeader,
                Tag: param.Tag,
                ColAsSerialize: param.ColAsSerialize
            };
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
    ExportXLS = function (options, guid) {
            if ($("#excelForm").length == 0) {
                $('<form id="excelForm"  method="post" target="excelIFrame"><input type="hidden" name="excelParam" id="excelData" /></form><iframe id="excelIFrame" name="excelIFrame" style="display:none;"></iframe>').appendTo("body");
            }
            if (options.FileFormat == "undefined") {
                options.FileFormat = "0";
            }
            if (!options.url) {
                options.url = "../Import/Export";
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
            //默认不分页导出数据
            if (!$.isArray(options.data)) {
                //查询条件
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

jQuery.extend(String.prototype,{
    /*
   * 功能：    类似C# String.Format()格式化功能
   * 参数：    args：参数
   * 返回值：  无
   * 创建人：  杜冬军
   */
    format: function (args) {
        var result = this;
        if (arguments.length > 0) {
            if (arguments.length == 1 && typeof (args) == "object") {
                for (var key in args) {
                    if (args[key] != undefined) {
                        var reg = new RegExp("\\{" + key + "\\}", "g");
                        result = result.replace(reg, args[key]);
                    }
                }
            }
            else {
                for (var i = 0; i < arguments.length; i++) {
                    if (arguments[i] != undefined) {
                        //var reg = new RegExp("({[" + i + "]})", "g");//这个在索引大于9时会有问题，谢谢何以笙箫的指出
                        var reg = new RegExp("\\{" + i + "\\}", "g");
                        result = result.replace(reg, arguments[i]);
                    }
                }
            }
        }
        return result;
    },
});
jQuery.extend({
    myParam: function (a, traditional) {
        var prefix, s = [], rbracket = /\[\]$/,
            add = function (key, value) {
                value = jQuery.isFunction(value) ? value() : (value == null ? "" : value);
                s[s.length] = encodeURIComponent(key) + "=" + encodeURIComponent(value);
            },
            buildParams = function (prefix, obj, traditional, add) {
                var name;
                if (jQuery.isArray(obj)) {
                    jQuery.each(obj, function (i, v) {
                        if (traditional || rbracket.test(prefix)) {
                            add(prefix, v);
                        } else {
                            // Item is non-scalar (array or object), encode its numeric index.
                            buildParams(prefix + "[" + (typeof v === "object" ? i : "") + "]", v, traditional, add);
                        }
                    });

                } else if (!traditional && jQuery.type(obj) === "object") {
                    // Serialize object item.
                    for (name in obj) {
                        buildParams(prefix + "[" + name + "]", obj[name], traditional, add);
                    }

                } else {
                    // Serialize scalar item.
                    add(prefix, obj);
                }
            };

        if (traditional === undefined) {
            traditional = jQuery.ajaxSettings && jQuery.ajaxSettings.traditional;
        }

        if (jQuery.isArray(a) || (a.jquery && !jQuery.isPlainObject(a))) {
            // Serialize the form elements
            jQuery.each(a, function () {
                add(this.name, this.value);
            });

        } else {
            for (prefix in a) {
                buildParams(prefix, a[prefix], traditional, add);
            }
        }

        // Return the resulting serialization
        return s.join("&");
    },
    download: function (url, data, method) {
        if (url && data) {
            method = method || 'post';
            data = typeof (data) == "string" ? data : decodeURIComponent($.myParam(data));
            var inputs = '';
            $.each(data.split('&'), function () {
                var pair = this.split('=');
                inputs += '<input type="hidden"   name="{0}" value="{1}"/>'.format(pair[0], pair[1]);
            });
            var objForm = $("#fileForm");
            if (objForm.length == 0) {
                objForm = $('<form id="fileForm" method="{0}" target="fileIFrame" action="{1}">{2}</form><iframe id="fileIFrame" name="fileIFrame" style="display:none;"></iframe>'.format(method, url, inputs)).appendTo('body');
            } else {
                objForm.attr("method", method).attr("action", url).html(inputs);
            }
            objForm.submit();
        }
    },
    DownloadExcelTemplate: function (type) {
        if (type == "undefined") {
            return;
        }
        var param = { type: type };
        $.download("../Import/DownLoadTemplate", param, "get");
    }
});