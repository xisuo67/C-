/*
 *说明:通用下拉框列表
 *作者:赵贤浩
 *日期:2016.05.01
 * 依赖 jquery.js, jquery.jqGrid.min.js,JSONQuery.js
 */
;(function ($) {

    $.fn.extend({
        //下拉框列表带搜索
        combogrid: function (options) {
            var defaluts = {
                datatype : "local",
                multiselect: false,// 是否单选
                searchfilter: null,
                afterSelected: null,//选中单行记录回调,参数 row(单选)或rows(多选)
                textField: "text",//必须，显示在文本框的文本字段
                valueField: "value",//必须，附加在input 元素的data-value 属性
                datatype: "json",
                idField: "id",//必须，唯一关键字段
                height: 300,
                width: 508,
                rowNum: 10,
                rowcontent: true,
                rowList: [5, 10, 20, 50],
                autowidth: true,
                rownumbers: false,
                styleUI: "Bootstrap",
                rownumWidth: 40,
                loadtext: "努力为您加载中...",
                loadui: "block",
                mtype: 'post',
                pager: "grid_Cross_Pages_" + id + "",
                pagerpos: "left",
                viewrecords: true,
                hidegrid: false,
                prmNames: { page: "page", rows: "pagesize", sort: "sortfield", order: "sortorder", search: null, nd: null, npage: null },
                jsonReader: {
                    repeatitems: false,
                    root: "Result",
                    total: "TotalPage",
                    records: "TotalCount"
                },
                emptyrecords: "",

                //行选择事件
                'onSelectRow': function (rowid, status) {
                    if (!opts.multiselect) {
                        radioSelect = rowid;
                        $("#grid_Cross_" + id + " tbody tr").removeClass("success").attr("aria-selected", "false").find("input[type=checkbox]").prop("checked", false);
                        $("#grid_Cross_" + id + " tbody tr[id=" + rowid + "]").addClass("success").attr("aria-selected", "true").find("input[type=checkbox]").prop("checked", true);
                        var row = JSONQuery("?id='{0}'".format(rowid), crossData)[0];
                        $(that).val(row[opts.textField]).data('value', row[opts.valueField]);
                        if ($.isFunction(opts.afterSelected)) {
                            opts.afterSelected.call(that, row);
                        }
                    } else {
                        selectRowHandle(rowid, status);
                        var texts = [], values = [];
                        var rows = JSONQuery("?IsSelected=true", crossData);
                        for (var i = 0, length = rows.length; i < length; i++) {
                            var  row = rows[i];
                            texts.push(row[opts.textField]);
                            values.push(row[opts.valueField]);
                        }
                        $(that).val(texts.join(',')).data('value', values.join(','));
                        if ($.isFunction(opts.afterSelected)) {
                            opts.afterSelected.call(that, rows);
                        }
                    }

                },
                //全选取消全选事件
                'onSelectAll': function (aRowids, status) {
                    for (var i = 0, length = aRowids.length; i < length; i++) {
                        selectRowHandle(aRowids[i], status);
                    }
                    //给文本框赋值
                    var arrval = getJglkValue();
                    var objVal = getObjVal();
                    var texts = [], values = [];
                    var rows = JSONQuery("?IsSelected=true", crossData);
                    for (var i = 0, length = rows.length; i < length; i++) {
                        var row = rows[i];
                        texts.push(row[opts.textField]);
                        values.push(row[opts.valueField]);
                    }

                    $(that).val(texts.join(',')).data('value', values.join(','));
                    if ($.isFunction(opts.afterSelected)) {
                        opts.afterSelected.call(that, rows);
                    }

                },
                'gridComplete': function () {
                    if (!opts.multiselect) {
                        $("#cb_grid_Cross_" + id + "").remove();
                    }
                    var obj = $(this);
                    var data = obj.getRowData();
                    $(data).each(function (i, item) {
                        if (item.IsSelected && item.IsSelected.toBool() == true) {
                            obj.setSelection(item.id, false);
                        }
                    });
                }

            };

            var gridCross = null, crossData = [], objGridCross = null, that = this, flag, radioSelect = null;
            var id = $(that).attr("id");
            var opts = $.extend({}, defaluts, options);
            var wrap = $('<div class="allCrossDiv" id="allCrossDiv_' + id + '"><div class="cross-search"><input type="text" placeholder="关键字/简拼/全拼检索" class="form-control" id="txt_searchkey_' + id + '"></div><table id="grid_Cross_' + id + '"></table><div id="grid_Cross_Pages_' + id + '"></div></div>');
            $("body").prepend(wrap);
            if (opts.pager) {
                opts.pager = "grid_Cross_Pages_" + id + "";
            }
            $(document).on('click', function (e) {
                if ($(that).prop("disabled")) {
                    return;
                }
                var idIndex = $(e.target).attr("id") ? $(e.target).attr("id").indexOf("txt_searchkey_") : -1;
                var isPageSelect = $(e.target).hasClass("ui-pg-selbox");
                if ($(e.target).get(0) == $(that).get(0)) {

                    if ($("#allCrossDiv_" + id + "").css("display") == "none") {

                        var selected = JSONQuery("?IsSelected=true", crossData);
                        var notselected = JSONQuery("?IsSelected=!true", crossData);
                        var newData = selected.concat(notselected);
                        $("#grid_Cross_" + id + "").jqGrid("clearGridData");
                        $("#grid_Cross_" + id + "").jqGrid('setGridParam', { "data": newData }).trigger('reloadGrid');
                      
                        $("#allCrossDiv_" + id + "").show();
                    } else {
                        $("#allCrossDiv_" + id + "").hide();
                    }
                } else if ($(e.target).closest("#allCrossDiv_" + id + "").length == 1 && opts.multiselect) {
                    $("#allCrossDiv_" + id + "").show();
                }
                else if ($(e.target).closest("#allCrossDiv_" + id + "").length == 1 && idIndex == 0) {
                    $("#allCrossDiv_" + id + "").show();

                }
                else if ($(e.target).closest("#allCrossDiv_" + id + "").length == 1 && isPageSelect) {
                    $("#allCrossDiv_" + id + "").show();
                }
                else {
                    $("#allCrossDiv_" + id + "").hide();
                }
            });

            function getJglkValue() {
                var rows = JSONQuery("?IsSelected=true", crossData);
                var arrValue = [];
                var values = "", texts = "", row;
                for (var i = 0, length = rows.length; i < length; i++) {
                    row = rows[i];
                    if (i < length - 1) {
                        texts += row[opts.textField] + ",";
                        values += row[opts.valueField] + ",";
                    } else {
                        texts += row[opts.textField];
                        values += row[opts.valueField];
                    }
                }
                arrValue.push(values);
                arrValue.push(texts);
                return arrValue;
            }

            //返回对象
            function getObjVal() {
                var rows = JSONQuery("?IsSelected=true", crossData);
                var arrValue = [];
                var row;
                for (var i = 0, length = rows.length; i < length; i++) {
                    var obj = {};
                    row = rows[i];
                    obj.id = row[opts.valueField];
                    obj.name = row[opts.textField];
                    arrValue.push(obj);
                }
                return arrValue;
            }

            //选择路口界面 行选择处理
            function selectRowHandle(rowid, status) {
                var row = JSONQuery("?id='{0}'".format(rowid), crossData);
                if (row[0].IsSelected != status) {
                    row[0].IsSelected = status;
                }
            }

            function initCrossGrid(data) {
                opts.datatype = "local";
                opts.data = data;
                opts.height = 300;
                $("#grid_Cross_" + id + "").jqGrid(opts);
                if (opts.width) {
                    $("#grid_Cross_" + id + "").setGridWidth(opts.width);
                }
               

            }

            $.ajax({
                type: opts.type,
                url: opts.url,
                data:opts.data||null,
                dataType: "json",
                success: function (data) {
                    var datas = data.Result;
                    $(datas).each(function (i, item) {
                        item.id = item[opts.idField];
                        item.IsSelected = false;
                    });
                    crossData = datas;
                    //初始化选择路口列表
                    initCrossGrid(datas);
                }
            });


            //查询
            $("#txt_searchkey_" + id + "").on("keyup", function () {
                clearTimeout(flag);
                //延时400ms执行请求事件，如果感觉时间长了，就用合适的时间
                //只要有输入则不执行keyup事件
                flag = setTimeout(function () {
                    searchCross();
                }, 400);
            });

            //查询
            function searchCross() {
                //匹配路口编号或者路口名称或者拼音
                var result = [];
                if ($.isFunction(opts.searchfilter)) {
                    result = opts.searchfilter.call(that, crossData, $("#txt_searchkey_" + id + "").val());
                    $("#grid_Cross_" + id + "").clearGridData();
                    $("#grid_Cross_" + id + "").setGridParam({ "data": result });
                    $("#grid_Cross_" + id + "").trigger("reloadGrid");
                }

            }
            return this.each(function () {
                var $this = $(this);
                $this.click(function () {
                    var offset = $this.offset();
                    var height = $(this).outerHeight();
                    wrap.css({ left: offset.left, top: offset.top + height + 5 });
                });
            });

        }

    });
})(jQuery);