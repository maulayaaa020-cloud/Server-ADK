<!DOCTYPE html>
<html>
<head>
    <title>Upload - ADK</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>

<div class="hero">
    <div class="hero-right">
        <div class="card">

            <!-- FIX DI SINI -->
            <form action="proses.php" method="POST" enctype="multipart/form-data">

                <label class="upload-box">
                    <p id="upload-text"><b>Pilih File</b></p>
                    <span>Format: Word (.docx)</span>
                    <input type="file" name="file" id="fileInput" class="file-input" required>
                </label>

                <div class="paket">
                    <p>Pilih Paket:</p>

                    <label class="paket-item">
                        <input type="radio" name="paket" value="paket1" required>
                        <span>Paket 1</span>
                    </label>

                    <label class="paket-item">
                        <input type="radio" name="paket" value="paket2">
                        <span>Paket 2</span>
                    </label>

                    <label class="paket-item">
                        <input type="radio" name="paket" value="paket3">
                        <span>Paket 3</span>
                    </label>

                    <label class="paket-item">
                        <input type="radio" name="paket" value="cepat">
                        <span>Cepat</span>
                    </label>
                </div>

                <button type="submit" class="btn-Upload">
                    Lanjut ke Pembayaran
                </button>

            </form>

        </div>
    </div>
</div>

<script>
const fileInput = document.getElementById("fileInput");
const text = document.getElementById("upload-text");

fileInput.addEventListener("change", () => {
    if (fileInput.files.length > 0) {
        text.innerText = fileInput.files[0].name;
    }
});
</script>

</body>
</html>