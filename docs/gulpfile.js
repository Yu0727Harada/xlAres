const gulp = require("gulp");
const ejs = require("gulp-ejs");
const rename =require("gulp-rename");
const sass = require("gulp-sass");

gulp.task("ejs",function () {
    return(
    gulp
        .src(
            ["page/**/*.ejs","!page/partical/**/*.ejs"])
        .pipe(ejs())
        .pipe(rename({extname:".html"}))
        .pipe(gulp.dest("document"))
    );

});

gulp.task("sass",function () {
    return(
    gulp
            .src(["page/*.scss"])
            .pipe(sass())
            .pipe(gulp.dest("document"))
    );

});