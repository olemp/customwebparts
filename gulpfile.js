var gulp = require("gulp"),
    ts = require("gulp-typescript"),
    watch = require('gulp-watch'),
    rename = require("gulp-rename"),
    uglify = require('gulp-uglify'),
    runSequence = require("run-sequence"),
    merge = require('merge2');

gulp.task('file-watch', function () {
    watch("./src/om.customwebparts.ts").on("change", function () {
        runSequence("ts-compile", "js-compress");
    });
});

gulp.task('ts-compile', function () {
    console.log("Building dist...");
    var tsResult = gulp.src("./src/om.customwebparts.ts")
        .pipe(ts({
            declaration: true,
            outFile: "om.customwebparts.js",
            removeComments: true
        }));
    return merge([
        tsResult.dts.pipe(gulp.dest('./dist')),
        tsResult.js.pipe(gulp.dest('./dist'))
    ]);
});

gulp.task('js-compress', function (cb) {
    console.log("Compressing dist...");
     gulp.src("./dist/om.customwebparts.js")
    .pipe(uglify())
    .pipe(rename({
      suffix: '.min'
    }))
         .pipe(gulp.dest('./dist'));
});