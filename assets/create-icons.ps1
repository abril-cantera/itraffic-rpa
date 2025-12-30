# Script para crear iconos PNG para el Office Add-in

Add-Type -AssemblyName System.Drawing

function Create-Icon {
    param(
        [int]$Size,
        [string]$OutputPath
    )
    
    # Crear bitmap
    $bitmap = New-Object System.Drawing.Bitmap($Size, $Size)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    
    # Fondo azul Microsoft
    $blueBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0, 120, 212))
    $graphics.FillRectangle($blueBrush, 0, 0, $Size, $Size)
    
    # Texto blanco (E de Email)
    $fontSize = [Math]::Floor($Size * 0.7)
    $font = New-Object System.Drawing.Font("Arial", $fontSize, [System.Drawing.FontStyle]::Bold)
    $whiteBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    
    $text = "E"
    $textSize = $graphics.MeasureString($text, $font)
    $x = ($Size - $textSize.Width) / 2
    $y = ($Size - $textSize.Height) / 2
    
    $graphics.DrawString($text, $font, $whiteBrush, $x, $y)
    
    # Guardar
    $bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
    
    # Limpiar
    $graphics.Dispose()
    $bitmap.Dispose()
    $font.Dispose()
    $blueBrush.Dispose()
    $whiteBrush.Dispose()
    
    Write-Host "[OK] Creado: $OutputPath (${Size}x${Size})" -ForegroundColor Green
}

Create-Icon -Size 16 -OutputPath "icon-16.png"
Create-Icon -Size 32 -OutputPath "icon-32.png"
Create-Icon -Size 64 -OutputPath "icon-64.png"
Create-Icon -Size 80 -OutputPath "icon-80.png"

Write-Host ""
Write-Host "Todos los iconos creados exitosamente" -ForegroundColor Green
