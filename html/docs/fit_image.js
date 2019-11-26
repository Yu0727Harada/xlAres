function image_view(){
    let ele = document.getElementsByClassName('img_small');
    for(let i=0;i<ele.length;i++){
        //ele[i].style.width = 'auto';
        var parent = ele[i].parentNode;
        var parent_width = parent.clientWidth;
        if(ele[i].naturalWidth >= parent_width){
            ele[i].style.width = parent_width+'px';
        }else{
            //ele[i].style.width = '500px';
            ele[i].style.width = ele[i].naturalWidth+'px';
        }
        ele[i].style.height = 'auto';
        ele[i].style.padding = '20px';
        ele[i].style.margin = 'auto';
    }


}

window.onload = function(){
    image_view();
};
window.onresize = function(){
    image_view();
};