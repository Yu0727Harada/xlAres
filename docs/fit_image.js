function image_view(){
    let ele_small = document.getElementsByClassName('img_small');
    for(let i=0;i<ele_small.length;i++){
        //ele_small[i].style.width = 'auto';
        var parent_small = ele_small[i].parentNode;
        var parent_width_small = parent_small.clientWidth;
        if(ele_small[i].naturalWidth >= parent_width_small){
            ele_small[i].style.width = parent_width_small+'px';
        }else{
            //ele_small[i].style.width = '500px';
            ele_small[i].style.width = ele_small[i].naturalWidth+'px';
        }
        ele_small[i].style.height = 'auto';
        ele_small[i].style.padding = '20px';
        ele_small[i].style.margin = 'auto';
    }

    let ele_max = document.getElementsByClassName('img_max');
    for(let i=0;i<ele_max.length;i++){
        //ele_max[i].style.width = 'auto';
        var parent_max = ele_max[i].parentNode;
        var parent_width_max = parent_max.clientWidth;
        if(ele_max[i].naturalWidth*0.3 >= parent_width_max){
            ele_max[i].style.width = parent_width_max+'px';
        }else{
            //ele_max[i].style.width = '500px';
            ele_max[i].style.width = ele_max[i].naturalWidth*0.3+'px';
        }
        ele_max[i].style.height = 'auto';
        ele_max[i].style.padding = '20px';
        ele_max[i].style.margin = 'auto';
    }



}

window.onload = function(){
    image_view();
};
window.onresize = function(){
    image_view();
};