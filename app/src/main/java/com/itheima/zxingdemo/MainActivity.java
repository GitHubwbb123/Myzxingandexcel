package com.itheima.zxingdemo;

import android.content.Context;
import android.content.Intent;
import android.graphics.Bitmap;
import android.os.Bundle;
import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.ImageView;
import android.widget.TextView;
import android.widget.Toast;
import com.google.zxing.activity.CaptureActivity;
import com.google.zxing.common.BitmapUtils;



import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class MainActivity extends AppCompatActivity implements View.OnClickListener {

    private Button mBtn1;
    private EditText mEt;
    private Button mBtn2;
    private ImageView mImage;
    private final static int REQ_CODE = 1028;
    private Context mContext;
    private TextView mTvResult;
    private ImageView mImageCallback;

    File file;
    Button writecontes;
    EditText row;
    EditText col;
    int rowNum;
    int colNum;
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        initView();
        file=new File(Environment.getExternalStorageDirectory().getPath()+"/Test");
        if(!file.exists()){
        file.mkdir();
        }
        file=new File(Environment.getExternalStorageDirectory().getPath()+"/Test/test.xls");
        if(!file.exists()){
            try {
            //file.createNewFile();注意，在建立excel后，要建立相应的表格即sheet，后面才能getSheet，不然写不进去，
               // 之前能建立Excel，但是写不进去，其实是ile.createNewFile()只新建了文件，没有新建sheet,
                // 在createExcel(file)函数中，新建了Excel，还新建了shette,所以可行了。
                createExcel(file);//在createExcel里新建excel，并新建sheet
            }
            catch (Exception e)
            {

            }
        }
        mContext = this;
    }
    public void createExcel(File file) {
        WritableWorkbook wookbook = null;
        try {
            file.createNewFile();
            //没必要用文件流
            wookbook = Workbook.createWorkbook(file);
            WritableSheet sheet = wookbook.createSheet("第一张表", 0);
           // Label lable1 = new Label(0, 0, "姓名");
            //Label lable2 = new Label(1, 0, "年龄");
            //Label lable3 = new Label(2, 0, "性别");
            //sheet.addCell(lable1);
            //sheet.addCell(lable2);
            //sheet.addCell(lable3);
            wookbook.write();
            wookbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {

        }

    }
    private void initView() {
        writecontes = (Button) findViewById(R.id.wt);
        writecontes.setOnClickListener(this);
        mBtn1 = (Button) findViewById(R.id.btn1);
        mBtn1.setOnClickListener(this);
        mEt = (EditText) findViewById(R.id.et);
        mBtn2 = (Button) findViewById(R.id.btn2);
        mBtn2.setOnClickListener(this);
        mImage = (ImageView) findViewById(R.id.image);
        mImage.setOnClickListener(this);
        mTvResult = (TextView) findViewById(R.id.tv_result);
        mTvResult.setOnClickListener(this);
        mImageCallback = (ImageView) findViewById(R.id.image_callback);
        mImageCallback.setOnClickListener(this);
    }

    @Override
    public void onClick(View v) {
        switch (v.getId()) {
            case R.id.btn1:
//                startActivity(new Intent(MainActivity.this, CaptureActivity.class));
                Intent intent = new Intent(mContext, CaptureActivity.class);
                startActivityForResult(intent, REQ_CODE);
                break;
            case R.id.btn2:
                mImage.setVisibility(View.VISIBLE);
                //隐藏扫码结果view
                mImageCallback.setVisibility(View.GONE);
                mTvResult.setVisibility(View.GONE);

                String content = mEt.getText().toString().trim();
                Bitmap bitmap = null;
                try {
                    bitmap = BitmapUtils.create2DCode(content);//根据内容生成二维码
                    mTvResult.setVisibility(View.GONE);
                    mImage.setImageBitmap(bitmap);
                } catch (Exception e) {
                    e.printStackTrace();
                }
                break;

            case R.id.wt:
                addResult( mTvResult.getText().toString(),file);
               break;
        }
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (requestCode == REQ_CODE) {
            mImage.setVisibility(View.GONE);
            mTvResult.setVisibility(View.VISIBLE);
            mImageCallback.setVisibility(View.VISIBLE);

            String result = data.getStringExtra(CaptureActivity.SCAN_QRCODE_RESULT);
            Bitmap bitmap = data.getParcelableExtra(CaptureActivity.SCAN_QRCODE_BITMAP);

            mTvResult.setText("扫码结果："+result);
            showToast("扫码结果：" + result);
            if(bitmap != null){
                mImageCallback.setImageBitmap(bitmap);//现实扫码图片
            }
        }


    }

    private void showToast(String msg) {
        Toast.makeText(mContext, "" + msg, Toast.LENGTH_SHORT).show();
    }





    public void addResult(String str,File file) {
        Workbook original = null;
        WritableWorkbook workbook = null;
        try {//  如果是想要修改一个已存在的excel工作簿，则需要先获得它的原始工作簿，再创建一个可读写的副本：
            original = Workbook.getWorkbook(file);
            workbook = Workbook.createWorkbook(file, original);
            WritableSheet sheet = workbook.getSheet(0);
            int row = sheet.getRows();//cpl，列，row,行
            Label label = new Label(0, row, str);
            sheet.addCell(label);
            workbook.write();
        } catch (Exception e) {
        } finally {
            if (original != null) {
                original.close();
            }
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                } catch (WriteException e) {
                    e.printStackTrace();
                }
            }

        }
    }
}


