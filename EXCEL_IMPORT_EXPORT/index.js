var express     = require('express');  
var mongoose    = require('mongoose');  
var multer      = require('multer');  
var path        = require('path');  
var user = require('./models/schema_demo')
var excelToJson = require('convert-excel-to-json');
var csvtojson = require('csvtojson')
var bodyParser  = require('body-parser');  
const Excel = require('exceljs');
const csv=require('csvtojson')
var path = require('path');
const { setDefaultResultOrder } = require('dns/promises');
var app = express();

                      /**  Mongodb connection   */

mongoose.connect('mongodb+srv://demo:ashish12345@cluster0.yy33d.mongodb.net/ExcelDemo',{useNewUrlParser:true})  
.then(()=>console.log('connected to db'))  
.catch((err)=>console.log(err))  
   
                       /**  Upload Images  */


app.use(bodyParser.urlencoded({extended:false}));    
app.use(express.static(path.resolve(__dirname,'public')));   
const storage = multer.diskStorage({
  destination: "./public",
  filename: (req, file, cb) => {
    return cb(
      null,
      `${file.fieldname}_${Date.now()}${path.extname(file.originalname)}`
    );
  },
});
const upload = multer({
  storage: storage,
  limits: {
    fileSize: 1024 * 1024 * 3,
  },
});
app.use("/profile", express.static("public"));


                      /** Inserted Data    */

app.post("/upload", upload.single("Profile"), async(req, res,) => {
    const Registation =await user({
        Name : req.body.Name,
        Email : req.body.Email,
        Password : req.body.Password,
        Profile: `http://localhost:3002/profile/${req.file.filename}`,
    })
    const DataSave =await Registation.save()
    console.log(DataSave);    
    if(DataSave){      
        res.status(200).json(Registation)
    }
    else{
        res.send("not registation")
    }     
})

                   /**    Data Exports in ExcelFile     */

app.get("/excel", async (req,res,next) => { 
try{  
  
    var workbook = new Excel.Workbook();   
    const sheet = workbook.addWorksheet("userModel", {
      properties: { tabColor: { argb: 'black' } },
    });
    sheet.columns = [
      { header: 'ID', key: 'id', width: 30 },
      { header: 'Name', key: 'Name' },
      { header: 'Email', key: 'Email' },
      { header: 'Password', key: 'Password' },  
      { header: 'Profile', key: 'Profile' },
    ];
    
    const userData = await user.find();
    
    console.log(userData);
    userData.forEach((user,pass) => {
             
   
      const id = user.id;
      const rw = sheet.addRow({
        id: id.replace(/'/g, `"`),
        Name: user.Name,
        Email: user.Email,
        Password: user.Password,
        Profile: user.Profile

      });    
      rw.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '90C9BF' },
      };
  for (let i = 0; i < length; i++) {
    worksheet.getColumn(`A${i}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"One,Two,Three,Four"']
    }

    
  }
    });
    res.attachment('userModel.xlsx');
    await workbook.xlsx.write(res);
} catch(err){
 
    res.send(err)
}  
});

                    /**      Import excelfile save Database     */


app.post('/uploadExcelFile1',upload.single("file"),async (req, res) =>{  
 
    const file = req.file.filename;
    const parsedData = [];
    var errorArry = [];

    await csvtojson()
      .fromFile(`public/${file}`) 
      .then((data) => {
        for (let i = 0; i < data.length; i++) {
          if (errorArry.length == 0) {
            const newUser = new user({
              Name: data[i]["Name"],
              Email: data[i]["Email"],
              Password: data[i]["Password"],
              Profile: data[i]["Profile"],
            });
            newUser.save();
            parsedData.push(newUser);
          }
        }
      });
      console.log("Data Importd successfully");
     res.send("Data Importd successfully"); 
})

               /**   connection port */

               
app.listen(3002,()=>console.log('server run at port '));  