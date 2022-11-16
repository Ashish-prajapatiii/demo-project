var mongoose  =  require('mongoose');  
   
var excelSchema = new mongoose.Schema({  
    Name:{  
        type:String  
    },  
    Email:{  
        type:String  
    },    
    Password:{  
        type:String  
    },
    Profile:{
        type:String
    }
});  
   
module.exports = mongoose.model('userModel',excelSchema);  