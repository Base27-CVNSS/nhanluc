function onClickUploadFile(id, isVaiTroCapNhatTrucTiep) {
    var idNhomFile = id.split('-')[1];
    openDialogUpload(idNhomFile, isVaiTroCapNhatTrucTiep);
}

function fileChange(id) {
    var idNhomFile = id.split('-')[1];
    var input = $('#' + id)[0];
    if (input.files.length > 0) {
        var check = true;
        var formData = new FormData();
        for (var i = 0; i < input.files.length; i++) {
            var fileName = input.files[i].name;
            var ext = fileName.substring(fileName.lastIndexOf('.') + 1).toLowerCase();
            if (ext == "doc" || ext == "docx" || ext == "xls" || ext == "xlsx" || ext == "ppt" || ext == "pptx" || ext == "zip" || ext == "rar" || ext == "jpeg" || ext == "png" || ext == "jpg" || ext == "pdf" || ext == "txt") {
                formData.append("fileUpload", input.files[i]);
            }
            else {
                check = false;
                break;
            }
        }

        if (!check) alert("Các định dạng hỗ trợ: pdf,txt,doc,docx,xls,xlsx,ppt,pptx,zip,rar,jpeg,png,jpg");
        else {
            var xmlhttp;
            if (window.XMLHttpRequest) {// code for IE7+, Firefox, Chrome, Opera, Safari
                xmlhttp = new XMLHttpRequest();
            }
            else {// code for IE6, IE5
                xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
            }
            var arrObjReturn = null;
            xmlhttp.onreadystatechange = function () {
                if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {
                    var returnResult = xmlhttp.responseText;
                    arrObjReturn = JSON.parse(returnResult);
                }
            }
            xmlhttp.open("POST", RoutingUtil.EncryptURL("/Pages/processCommandAjax.aspx?Command=UploadFileTaiLieuHop"), false);
            xmlhttp.send(formData);
            //ndtai add
            if (arrObjReturn != "undefined" && arrObjReturn.length > 0 && arrObjReturn[0]["VuotNguongDungLuong"] == "true") {
                alert("Không thể tải lên do đã vượt ngưỡng dung lượng lưu trữ");
                return;
            }
            //ndtai endadd
            renderNewDanhSachFile(idNhomFile, arrObjReturn);
        }
    }
}


function uploadTaiLieu(formData, idNhomFile, isReload) {
    $.ajax({
        url: RoutingUtil.EncryptURL("/Pages/processCommandAjax.aspx?Command=UploadTaiLieuHop"),
        type: 'post',
        data: formData,
        processData: false,
        contentType: false,
        success: (function (data) {
            //renderNewDanhSachFile(idNhomFile, data);
            if (isReload) closeDialogAndRefresh("themtailieuhop-popup");
            else closeDialog("themtailieuhop-popup");
        })
    });
}

function openDialogUpload(idNhomFile, isVaiTroCapNhatTrucTiep) {
    var urlDialog = RoutingUtil.EncryptURL('/nqrk93gnlnrG0t1lUXSfkYUeF0scvQiSsNW4smzY72iMbjOekb6NQhJjKN3ltjs5AaSK9tF5LESGnfro93iLFQ_3D_3D?idNhomFile=' + idNhomFile + '&isVaiTroCapNhatTrucTiep=' + isVaiTroCapNhatTrucTiep + '&isReload=true&actionForm=Frontend');
    openDialog('themtailieuhop-popup', urlDialog);
}

function checkButtonViTri(idNhomFile) {
    var length = $("#list_btnFileUpload-" + idNhomFile + " .row").length;
    if (length == 1) {
        $(".upRow").hide();
        $(".downRow").hide();
    }
    else {
        $("#list_btnFileUpload-" + idNhomFile + " .row").each(function () {
            if ($(this).index() == 0) {
                $(this).find("div:eq(1)").find(".upRow").hide();
                $(this).find("div:eq(1)").find(".downRow").show();
            }
            else if ($(this).index() == length - 1) {
                $(this).find("div:eq(1)").find(".upRow").show();
                $(this).find("div:eq(1)").find(".downRow").hide();
            }
            else {
                $(this).find("div:eq(1)").find(".upRow").show();
                $(this).find("div:eq(1)").find(".downRow").show();
            }
        });
    }
}

function renderNewDanhSachFile(idNhomFile, jsonObject) {
    var idsFileUpload = $('input[id*="hdfIDsFileUpload_' + idNhomFile + '"]').val();
    if (idsFileUpload == "") idsFileUpload = ";";

    var insertContent = "";
    if (jsonObject.length > 0) {
        for (i = 0; i < jsonObject.length; i++) {
            var objFiles = jsonObject[i];
            insertContent +=
                "<div id=\"rowFile_" + objFiles.id + "\" class=\"row\" style='margin-bottom: 5px;'>" +
                "<div class=\"col-sm-9 col-xs-9\">" +
                "<input type=\"text\" id=\"idInputTenFile_" + objFiles.id + "\" class=\"form-control\" placeholder=\"Nhập tên file\" value=\"" + objFiles.ten + "\" disabled>" +
                "</div>" +
                "<div class=\"col-sm-3 col-xs-3\">" +
                "<a class=\"btn-success btn edit\" href='#' onclick=\"return editFiles('" + objFiles.id + "')\">" +
                "<span class=\"fa fa-edit\"></span>" +
                "</a>&nbsp;" +
                "<a class=\"btn-success btn\" href='#' onclick=\"return xoaFiles('" + objFiles.id + "', '" + idNhomFile + "')\">" +
                "<span class=\"fa fa-trash\"></span>" +
                "</a>&nbsp;";

            //insertContent += "<a class=\"btn-success btn upRow\" href='#' onclick=\"return upViTriFiles(this,'" + objFiles.id + "', '" + idNhomFile + "')\">" +
            //    "<span class=\"fa fa-arrow-up\"></span>" +
            //    "</a>&nbsp;" +
            //    "<a class=\"btn-success btn downRow\" href='#' onclick=\"return downViTriFiles(this,'" + objFiles.id + "', '" + idNhomFile + "')\">" +
            //    "<span class=\"fa fa-arrow-down\"></span>" +
            //    "</a>";
            insertContent += "</div></div>";
            idsFileUpload += objFiles.id + ";";
        }
    }

    //$('#list_btnFileUpload-' + idNhomFile).append(insertContent);
    $('#list_btnFileUploadPaging-' + idNhomFile).append(insertContent);

    $('input[id*="hdfIDsFileUpload_' + idNhomFile + '"]').val(idsFileUpload);
    checkButtonViTri(idNhomFile);
}

function editFiles(idFile) {
    var urlDialog = RoutingUtil.EncryptURL('/nqrk93gnlnrG0t1lUXSfkYUeF0scvQiSsNW4smzY72iMbjOekb6NQhJjKN3ltjs5AaSK9tF5LESGnfro93iLFQ_3D_3D?ID=' + idFile + '&actionForm=Frontend');
    openDialog('themtailieuhop-popup', urlDialog);
}

function xoaFiles(idFile, idNhomFile) {
    var r = confirm("Bạn có muốn xóa tài liệu này?.");
    if (r) {
        $.ajax({
            url: RoutingUtil.EncryptURL('/Pages/processCommandAjax.aspx?Command=XoaTaiLieuHop'),
            type: 'post',
            dataType: 'json',
            data: {
                "id": idFile,
            },
            async: false,
            success: (function (data) {
                //console.log(data);
                $("div#rowFile_" + idFile).remove();
                idsFileUpload = $('input[id*="hdfIDsFileUpload"]').val();
                idsFileUpload = idsFileUpload.replace(";" + idFile + ";", ";");
                $('input[id*="hdfIDsFileUpload_' + idNhomFile + '"]').val(idsFileUpload);
            })
        });
        checkButtonViTri(idNhomFile);
        return false;
    }    
}

function upViTriFiles(element, idFile, idNhomFile) {
    index = $(element).parent('div').parent('div.row').index();
    if (index > 0) {
        idFilePrev = $("#rowFile_" + idFile).parent("div").find(".row:eq(" + (index - 1) + ")").attr("id").split('_')[1];

        src = $("#rowFile_" + idFilePrev).html();
        $("#rowFile_" + idFilePrev).html($("#rowFile_" + idFile).html());
        $("#rowFile_" + idFile).html(src);
        tempElement = $("#rowFile_" + idFile);
        $("#rowFile_" + idFilePrev).attr("id", "rowFile_" + idFile);
        tempElement.attr("id", "rowFile_" + idFilePrev);

        idsFileUpload = $('input[id*="hdfIDsFileUpload"]').val();
        idsFileUpload = idsFileUpload.replace(";" + idFile + ";", ";" + idFilePrev + ";");
        idsFileUpload = idsFileUpload.replace(";" + idFilePrev + ";", ";" + idFile + ";");
        $('input[id*="hdfIDsFileUpload_' + idNhomFile + '"]').val(idsFileUpload);
    }
    checkButtonViTri(idNhomFile);
    return false;
}

function downViTriFiles(element, idFile, idNhomFile) {
    index = $(element).parent('div').parent('div.row').index();
    length = $("#rowFile_" + idFile).parent('div').find('div.row').length;
    if (index < length - 1) {
        idFileNext = $("#rowFile_" + idFile).parent("div").find(".row:eq(" + (index + 1) + ")").attr("id").split('_')[1];

        src = $("#rowFile_" + idFileNext).html();
        $("#rowFile_" + idFileNext).html($("#rowFile_" + idFile).html());
        $("#rowFile_" + idFile).html(src);
        tempElement = $("#rowFile_" + idFile);
        $("#rowFile_" + idFileNext).attr("id", "rowFile_" + idFile);
        tempElement.attr("id", "rowFile_" + idFileNext);

        idsFileUpload = $('input[id*="hdfIDsFileUpload"]').val();
        idsFileUpload = idsFileUpload.replace(";" + idFile + ";", ";" + idFileNext + ";");
        idsFileUpload = idsFileUpload.replace(";" + idFileNext + ";", ";" + idFile + ";");
        $('input[id*="hdfIDsFileUpload_' + idNhomFile + '"]').val(idsFileUpload);
    }
    checkButtonViTri(idNhomFile);
    return false;
}

function openDialogSettingFiles(idFile) {
    var data = $("#txtTrangThaiTaiLieu_" + idFile).val().split('|');
    var truycap = $("#txtTrangThaiTruyCap_" + idFile).val().split('|');

    openDialog('chonthietlaptailieumat-popup', RoutingUtil.EncryptURL('/nqrk93gnlnrG0t1lUXSfkeSLsmJ4pDKH1Lb_2BJ1HIn1KIVp8hCfMTWNFcKZGRDlqj0j_2FRQzWBj65fsqYXTiIL4Q_3D_3D?idFile=' + idFile + '&trangThai=' + data[0]
        + '&trangThaiTaiLieuMat=' + data[1] + '&thoiGianTruyCap=' + data[2] + '&stt=0&idsNhanVienTruyCap=' + truycap[0] + "&idsDonViTruyCap=" + truycap[2]));
}

function settingFiles(stt, idFile, trangThai, trangThaiTaiLieuMat, thoiGianTruyCap, listNhanVienTruyCap, listDonViTruyCap) {
    var buttonSetting = $("#rowFile_" + idFile).find('a.cog');
    if (trangThaiTaiLieuMat == "HoatDong" || listNhanVienTruyCap.length >= 3 || listDonViTruyCap.length >= 3) {
        buttonSetting.removeClass('btn-success');
        buttonSetting.addClass('btn-primary');
    } else {
        buttonSetting.removeClass('btn-primary');
        buttonSetting.addClass('btn-success');
    }

    var truycap = $("#txtTrangThaiTruyCap_" + idFile).val().split('|');
    $("#txtTrangThaiTaiLieu_" + idFile).val(trangThai + "|" + trangThaiTaiLieuMat + "|" + thoiGianTruyCap);
    $("#txtTrangThaiTruyCap_" + idFile).val(listNhanVienTruyCap + "|" + truycap[1] + "|" + listDonViTruyCap + "|" + truycap[3]);
}

function bindThongTinFile(idFile, idNhomFile, trangThai) {
    var showInfo = true;
    if ($('#list_btnFileUpload-' + idNhomFile).html().trim() != "") {
        if ($('#rowFile_' + idFile).length) {
            showInfo = false;
        }
        $('#list_btnFileUpload-' + idNhomFile).html("");
    }
    if (showInfo) {
        $.ajax({
            url: RoutingUtil.EncryptURL('/_2F8bqqeFGuX3P5IBV5ifYeh3_2FY_2B2VgMp_2FEVccR5fPLHJYUV03c6VWmfo_2FtQVSHyVw?Command=getListTaiLieuHopInfo'),
            type: 'post',
            dataType: 'html',
            data: {
                "listIDFiles": idFile,
            },
            success: (function (html) {
                var jsonObject = JSON.parse(html);
                //console.log(html);
                for (var i = 0; i < jsonObject.length; i++) {
                    classNameTaiLieuMat = (jsonObject[i].TrangThaiTaiLieuMat == 'HoatDong' || jsonObject[i].IDsNhanVienTruyCap.length >= 3 || jsonObject[i].IDsDonViTruyCap.length >= 3) ? "btn-primary" : "btn-success";
                    var insertContent =
                        "<div id=\"rowFile_" + jsonObject[i].ID + "\" class=\"row\" style='margin-bottom: 5px;'>" +
                        "<div class=\"col-sm-9 col-xs-9\">" +
                        "<input type=\"text\" id=\"idInputTenFile_" + jsonObject[i].ID + "\" class=\"form-control\" placeholder=\"Nhập tên file\" value=\"" + jsonObject[i].Ten + "\" disabled>" +
                        "</div>" +
                        "<div class=\"col-sm-3 col-xs-3\">" +
                        "<a class=\"btn-success btn edit\" href='#' onclick=\"return editFiles('" + jsonObject[i].ID + "')\">" +
                        "<span class=\"fa fa-edit\"></span>" +
                        "</a>&nbsp;" +
                        "<a class=\"" + classNameTaiLieuMat + " btn cog\" href='#' onclick=\"return openDialogSettingFiles('" + jsonObject[i].ID + "')\">" +
                        "<span class=\"fa fa-cog\"></span><input type='hidden' id='txtTrangThaiTaiLieu_" + jsonObject[i].ID + "' value='" + jsonObject[i].TrangThai + "|" + jsonObject[i].TrangThaiTaiLieuMat + "|" + jsonObject[i].ThoiGianTruyCap + "' />" +
                        "<input type='hidden' id='txtTrangThaiTruyCap_" + jsonObject[i].ID + "' value='" + jsonObject[i].IDsNhanVienTruyCap + "|" + jsonObject[i].IDsNhanVienTruyCap + "|" + jsonObject[i].IDsDonViTruyCap + "|" + jsonObject[i].IDsDonViTruyCap + "' />" +
                        "</a>";
                    //debugger;
                    //if (trangThai != "CongBo" && trangThai != "DaDuyet" && trangThai !) {
                    //    insertContent += "&nbsp;<a class=\"btn-success btn\" href='#' onclick=\"return xoaFiles('" + jsonObject[i].ID + "', '" + idNhomFile + "')\">" +
                    //        "<span class=\"fa fa-trash\"></span></a>";
                    //}
                    insertContent += "</div></div>";

                    $('#list_btnFileUpload-' + idNhomFile).append(insertContent);
                }
                checkButtonViTri(idNhomFile);
            })
        });
    }
}

function bindAllDanhSachFile(idNhomFile) {
    var listId = $('input[id*="hdfIDsFileUpload_' + idNhomFile + '"]').val();
    $('#list_btnFileUpload-' + idNhomFile).html("");
    if (listId.length >= 3) {
        $.ajax({
            url: RoutingUtil.EncryptURL('/_2F8bqqeFGuX3P5IBV5ifYeh3_2FY_2B2VgMp_2FEVccR5fPLHJYUV03c6VWmfo_2FtQVSHyVw?Command=getListFilePhienHopInfo'),
            type: 'post',
            dataType: 'html',
            data: {
                "listIDFiles": listId,
            },
            success: (function (html) {
                var jsonObject = JSON.parse(html);
                //console.log(html);
                for (var i = 0; i < jsonObject.length; i++) {
                    classNameTaiLieuMat = (jsonObject[i].TrangThaiTaiLieuMat == 'HoatDong' || jsonObject[i].IDsNhanVienTruyCap.length >= 3 || jsonObject[i].IDsDonViTruyCap.length >= 3) ? "btn-primary" : "btn-success";
                    var insertContent =
                        "<div id=\"rowFile_" + jsonObject[i].ID + "\" class=\"row\" style='margin-bottom: 5px;'>" +
                        "<div class=\"col-sm-9 col-xs-9\">" +
                        "<input type=\"text\" id=\"idInputTenFile_" + jsonObject[i].ID + "\" class=\"form-control\" placeholder=\"Nhập tên file\" value=\"" + jsonObject[i].Ten + "\" disabled>" +
                        "</div>" +
                        "<div class=\"col-sm-3 col-xs-3\">" +
                        "<a class=\"btn-success btn edit\" href='#' onclick=\"return editFiles('" + jsonObject[i].ID + "')\">" +
                        "<span class=\"fa fa-edit\"></span>" +
                        "</a>&nbsp;" +
                        "<a class=\"" + classNameTaiLieuMat + " btn cog\" href='#' onclick=\"return openDialogSettingFiles('" + jsonObject[i].ID + "')\">" +
                        "<span class=\"fa fa-cog\"></span><input type='hidden' id='txtTrangThaiTaiLieu_" + jsonObject[i].ID + "' value='" + jsonObject[i].TrangThai + "|" + jsonObject[i].TrangThaiTaiLieuMat + "|" + jsonObject[i].ThoiGianTruyCap + "' />" +
                        "<input type='hidden' id='txtTrangThaiTruyCap_" + jsonObject[i].ID + "' value='" + jsonObject[i].IDsNhanVienTruyCap + "|" + jsonObject[i].IDsNhanVienTruyCap + "|" + jsonObject[i].IDsDonViTruyCap + "|" + jsonObject[i].IDsDonViTruyCap + "' />" +
                        "</a>&nbsp;" +
                        "<a class=\"btn-success btn\" href='#' onclick=\"return xoaFiles('" + jsonObject[i].ID + "', '" + idNhomFile + "')\">" +
                        "<span class=\"fa fa-trash\"></span>" +
                        "</a>&nbsp;" +
                        "<a class=\"btn-success btn upRow\" href='#' onclick=\"return upViTriFiles(this,'" + jsonObject[i].ID + "', '" + idNhomFile + "')\">" +
                        "<span class=\"fa fa-arrow-up\"></span>" +
                        "</a>&nbsp;" +
                        "<a class=\"btn-success btn downRow\" href='#' onclick=\"return downViTriFiles(this,'" + jsonObject[i].ID + "', '" + idNhomFile + "')\">" +
                        "<span class=\"fa fa-arrow-down\"></span>" +
                        "</a>";
                    insertContent += "</div></div>";

                    $('#list_btnFileUpload-' + idNhomFile).append(insertContent);
                }
                checkButtonViTri(idNhomFile);
            })
        });
    }
}