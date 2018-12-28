package com.larapin.poiandroid

import android.support.v7.app.AppCompatActivity
import android.os.Bundle
import android.os.Environment
import android.widget.Button
import android.widget.EditText
import org.apache.poi.xslf.usermodel.SlideLayout
import org.apache.poi.xslf.usermodel.XMLSlideShow
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.jetbrains.anko.find
import org.jetbrains.anko.toast
import java.io.File
import java.io.FileOutputStream

class MainActivity : AppCompatActivity() {

    private lateinit var btnDocx: Button
    private lateinit var btnXlsx: Button
    private lateinit var btnPptx: Button

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        System.setProperty("org.apache.poi.javax.xml.stream.XMLInputFactory", "com.fasterxml.aalto.stax.InputFactoryImpl")
        System.setProperty("org.apache.poi.javax.xml.stream.XMLOutputFactory", "com.fasterxml.aalto.stax.OutputFactoryImpl")
        System.setProperty("org.apache.poi.javax.xml.stream.XMLEventFactory", "com.fasterxml.aalto.stax.EventFactoryImpl")
        
        val inputText = find(R.id.input_text) as EditText
        btnDocx = find(R.id.btn_docx)
        btnXlsx = find(R.id.btn_xlsx)
        btnPptx = find(R.id.btn_pptx)

        val path = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS)

        btnDocx.setOnClickListener {
            createDocx(path, inputText.text.toString().trim())
        }
        btnXlsx.setOnClickListener{
            createXlsx(path, inputText.text.toString().trim())
        }
        btnPptx.setOnClickListener{
            createPptx(path, inputText.text.toString().trim())
        }
    }

    private fun createDocx(path: File, message: String) {
        try {
            val document = XWPFDocument()

            val outputStream = FileOutputStream(File(path,"/poi.docx"))

            val paragraph = document.createParagraph()
            val run = paragraph.createRun()
            run.setText(message)

            document.write(outputStream)
            outputStream.close()
            toast("poi.docx was successfully created")
        }catch (e: Exception){
            e.printStackTrace()
        }
    }

    private fun createXlsx(path: File, message: String) {
        try {
            val workbook = XSSFWorkbook()

            val outputStream = FileOutputStream(File(path, "/poi.xlsx"))

            val sheet = workbook.createSheet("Sheet 1")
            val row = sheet.createRow(2)
            val cell = row.createCell(1)
            cell.setCellValue(message)

            workbook.write(outputStream)
            outputStream.close()
            toast("poi.xlsx was successfully created")
        }catch (e: Exception){
            e.printStackTrace()
        }
    }

    private fun createPptx(path: File, message: String) {
        try {
            val slideShow = XMLSlideShow()

            val outputStream = FileOutputStream(File(path, "/poi.pptx"))

            val slideMaster = slideShow.slideMasters[0]
            val titleLayout = slideMaster.getLayout(SlideLayout.TITLE)
            val slide = slideShow.createSlide(titleLayout)
            val title = slide.getPlaceholder(0)
            title.text = message

            slideShow.write(outputStream)
            outputStream.close()
            toast("poi.pptx was successfully created")
        }catch (e: Exception){
            e.printStackTrace()
        }
    }
}
