$(document).ready(function() {
    /*颜色随机*/
    var tags_a = $("#tags").find("a");
    tags_a.each(function() {
        var x = 9;
        var y = 0;
        var rand = parseInt(Math.random() * (x - y + 1) + y);
        $(this).addClass("size" + rand);
    });
});