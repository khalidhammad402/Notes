const socket = io();

socket.on('fileChange', ({data}) => {
    console.log(data);
    var content = '<div class="card"><h5 class="card-header">\
                '+data[0]+'</h5><div class="card-body">\
                <p class="card-text">Qualification : '+ data[1]+'</p>\
                <p class="card-text">Branch : '+ data[2]+'</p>\
                <p class="card-text">Date of Application : '+ data[3]+'</p>\
                <p class="card-text">State : '+ data[4]+'</p>\
                <a href="#" class="btn btn-primary">Go somewhere</a>\
                </div></div>'
    $("#12").append(content);
});
