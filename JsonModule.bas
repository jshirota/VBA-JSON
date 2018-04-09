Attribute VB_Name = "JsonModule"
Option Explicit

Private script As Object

Function Json(text As String, selector As String) As Variant
    
    If script Is Nothing Then
        
        ' Loads via Microsoft HTML Application host if necessary.
        Set script = CreateObjectx86("MSScriptControl.ScriptControl")
        script.Language = "JScript"
        
        ' Douglas Crockford
        script.AddCode ("""object""!=typeof JSON&&(JSON={}),function(){""use strict"";var rx_one=/^[\],:{}\s]*$/,rx_two=/\\(?:[""\\\/bfnrt]|u[0-9a-fA-F]{4})/g,rx_three=/""[^""\\\n\r]*""|true|false|null|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?/g,rx_four=/(?:^|:|,)(?:\s*\[)+/g,rx_escapabl" & _
        "e=/[\\""\u0000-\u001f\u007f-\u009f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,rx_dangerous=/[\u0000\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,ga" & _
        "p,indent,meta,rep;function f(t){return t<10?""0""+t:t}function this_value(){return this.valueOf()}function quote(t){return rx_escapable.lastIndex=0,rx_escapable.test(t)?'""'+t.replace(rx_escapable,function(t){var e=meta[t];return""string""==typeof e?e:""\" & _
        "\u""+(""0000""+t.charCodeAt(0).toString(16)).slice(-4)})+'""':'""'+t+'""'}function str(t,e){var r,n,o,u,f,a=gap,i=e[t];switch(i&&""object""==typeof i&&""function""==typeof i.toJSON&&(i=i.toJSON(t)),""function""==typeof rep&&(i=rep.call(e,t,i)),typeof i){c" & _
        "ase""string"":return quote(i);case""number"":return isFinite(i)?String(i):""null"";case""boolean"":case""null"":return String(i);case""object"":if(!i)return""null"";if(gap+=indent,f=[],""[object Array]""===Object.prototype.toString.apply(i)){for(u=i.lengt" & _
        "h,r=0;r<u;r+=1)f[r]=str(r,i)||""null"";return o=0===f.length?""[]"":gap?""[\n""+gap+f.join("",\n""+gap)+""\n""+a+""]"":""[""+f.join("","")+""]"",gap=a,o}if(rep&&""object""==typeof rep)for(u=rep.length,r=0;r<u;r+=1)""string""==typeof rep[r]&&(o=str(n=rep[r" & _
        "],i))&&f.push(quote(n)+(gap?"": "":"":"")+o);else for(n in i)Object.prototype.hasOwnProperty.call(i,n)&&(o=str(n,i))&&f.push(quote(n)+(gap?"": "":"":"")+o);return o=0===f.length?""{}"":gap?""{\n""+gap+f.join("",\n""+gap)+""\n""+a+""}"":""{""+f.join("","")" & _
        "+""}"",gap=a,o}}""function""!=typeof Date.prototype.toJSON&&(Date.prototype.toJSON=function(){return isFinite(this.valueOf())?this.getUTCFullYear()+""-""+f(this.getUTCMonth()+1)+""-""+f(this.getUTCDate())+""T""+f(this.getUTCHours())+"":""+f(this.getUTCMin" & _
        "utes())+"":""+f(this.getUTCSeconds())+""Z"":null},Boolean.prototype.toJSON=this_value,Number.prototype.toJSON=this_value,String.prototype.toJSON=this_value),""function""!=typeof JSON.stringify&&(meta={""\b"":""\\b"",""\t"":""\\t"",""\n"":""\\n"",""\f"":""" & _
        "\\f"",""\r"":""\\r"",'""':'\\""',""\\"":""\\\\""},JSON.stringify=function(t,e,r){var n;if(gap="""",indent="""",""number""==typeof r)for(n=0;n<r;n+=1)indent+="" "";else""string""==typeof r&&(indent=r);if(rep=e,e&&""function""!=typeof e&&(""object""!=typeof" & _
        " e||""number""!=typeof e.length))throw new Error(""JSON.stringify"");return str("""",{"""":t})}),""function""!=typeof JSON.parse&&(JSON.parse=function(text,reviver){var j;function walk(t,e){var r,n,o=t[e];if(o&&""object""==typeof o)for(r in o)Object.proto" & _
        "type.hasOwnProperty.call(o,r)&&(void 0!==(n=walk(o,r))?o[r]=n:delete o[r]);return reviver.call(t,e,o)}if(text=String(text),rx_dangerous.lastIndex=0,rx_dangerous.test(text)&&(text=text.replace(rx_dangerous,function(t){return""\\u""+(""0000""+t.charCodeAt(0" & _
        ").toString(16)).slice(-4)})),rx_one.test(text.replace(rx_two,""@"").replace(rx_three,""]"").replace(rx_four,"""")))return j=eval(""(""+text+"")""),""function""==typeof reviver?walk({"""":j},""""):j;throw new SyntaxError(""JSON.parse"")})}();")

    End If
    
    script.AddCode ("function select(text){ var o = JSON.parse(text)" & selector & "; return o == null ? '' : (typeof o === 'object' ? JSON.stringify(o) : o); }")

    Json = script.Run("select", text)
     
End Function
