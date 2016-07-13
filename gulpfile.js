var gulp = require("gulp"),
    tsc = require("gulp-typescript"),
    watch = require('gulp-watch');

gulp.task('file-watch', function () {
    watch("./src/om.customwebparts.ts").on("change", function (filePath) {
        console.log("Building dist...");
        var outFile = filePath;
        outFile = outFile.replace(".ts", ".js");
        outFile = outFile.replace("\\src\\", "\\dist\\");
        gulp
            .src(filePath)
            .pipe(tsc({
                outFile: outFile
            }))
            .pipe(gulp.dest("./"));
        gulp
            .src(filePath)
            .pipe(tsc({
                declaration: true
            }))
            .pipe(gulp.dest("./"));
    });
});