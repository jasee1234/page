<!DOCTYPE html>
<html>
<body>

<?php
$target_dir = "uploads/FCL/";
$target_file = $target_dir . basename($_FILES["fileToUpload"]["name"]);
$uploadOk = 1;
$imageFileType = pathinfo($target_file,PATHINFO_EXTENSION);



if ($uploadOk == 0) {
    echo "Sorry, your file was not uploaded.";

}
else {
    if (move_uploaded_file($_FILES["fileToUpload"]["tmp_name"], $target_file)) {
        echo "The file ". basename( $_FILES["fileToUpload"]["name"]). " has been uploaded.";
    } else {
        echo "Sorry, there was an error uploading your file.";
    }
}
?>

<form method="get" action="/index.html">
    <button type="submit">GO BACK</button>
</form>

</body>
</html>
