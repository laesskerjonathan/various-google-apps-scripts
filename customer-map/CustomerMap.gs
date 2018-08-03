//Before Run: See TODO
//Saved in root as map.png

function pxlsToKm(pxls) {
  return pxls * 430 / 1920;
}

function getIconDistance(size1, size2) {
  return (size1 + size2) / 2;
}

//jeweils anpassen
function getSize(amount){
  var sizes = [16, 24, 32, 48, 64];
  if(amount <= 5){
    return sizes[amount - 1];
  }
  return sizes[4];
}

function getDistance(c1, c2) {
  var lat1 = rad(c1.lat), lat2 = rad(c2.lat);
  var lng1 = rad(c1.lng), lng2 = rad(c2.lng);
  var dLng = (lng2-lng1), dLat = (lat2-lat1);
  var R = 6371/1.6;
  
  var a = Math.sin(dLat/2) * Math.sin(dLat/2) + 
    Math.sin(dLng/2) * Math.sin(dLng/2) * 
    Math.cos(lat1) *  Math.cos(lat2); 
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a)); 
      
  return parseInt(1.60934 * R * c);
}


function rad(degrees) {
  return degrees * Math.PI/180;
}


function generateMap() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
   // Add markers for the nearbye train stations.
  //map.setCustomMarkerStyle('icon:https://upload.wikimedia.org/wikipedia/commons/d/d5/Japan_small_icon.png', false);
  
  var plzJson = {};
  
  for (var i = 0; i < data.length; i++) {
    var plz = data[i][0].toString();
    if(plz !== "") {
      var coordinates = Maps.newGeocoder().geocode(plz + ", Schweiz");
      Utilities.sleep(100);
      var lng = coordinates.results[0].geometry.location.lng;
      var lat = coordinates.results[0].geometry.location.lat;
      
      if(typeof(plzJson[plz]) === "undefined"){
        plzJson[plz] = { 
          amount : 1,
          lng : lng,
          lat : lat,
          plz : plz,
          connections : []
        };
      } else {
        plzJson[plz].amount++;
      }
    }
  }

  //JSON in Array umwandeln
  var plzArray = [];
  for (var entry in plzJson){
    plzArray.push(plzJson[entry]);
  }
  
  //Jedes Element mit jedem anderen verknüpfen und Distanz überprüfen
  for(var i = 0; i < plzArray.length; i++) {
    plzArray[i].id = i;
    for(var j = i + 1; j < plzArray.length; j++) {
      var distance = getDistance(plzArray[i], plzArray[j]);
      var iconDistance = getIconDistance(getSize(plzArray[i].amount), getSize(plzArray[j].amount));
      
      if(pxlsToKm(iconDistance) > distance){
          plzArray[i].connections.push(j);
          plzArray[j].connections.push(i);
          
        //Logger.log("Überschneidung First: "); //+ plzArray[i].plz + ", Second: " + plzArray[j].plz);
      }
    }
  }
          
  plzArray.sort(function(a, b){return b.connections.length - a.connections.length;});
  
  //Logger.log(plzArray);
  
 printMap(plzArray, "oldMap");
  //Überschneidungen berechnen
  var newPlaces = [];
  
  for(var i = 0; i < plzArray.length; i++) {
    var newAmount = plzArray[i].amount;
    var newLng = plzArray[i].lng;
    var newLat = plzArray[i].lat;
    var newPlz = plzArray[i].plz;
    for(var k = i + 1; k < plzArray.length; k++) {
      for(var j = 0; j < plzArray[i].connections.length; j++){
        if(plzArray[k].id == plzArray[i].connections[j] && typeof(plzArray[k].alreadyUsed) === "undefined") {
          //Logger.log("2. Überschneidung: " + plzArray[k].plz + " 3: " + plzArray[i].plz);
          newAmount += plzArray[k].amount;
          newLng += plzArray[k].lng;
          newLat += plzArray[k].lat;
          plzArray[k].alreadyUsed = true;
          newPlz = "";
        }
      }
    }
    if(typeof(plzArray[i].alreadyUsed) === "undefined"){
      if(plzArray[i].connections.length > 0){
        newLng = newLng / newAmount;
        newLat = newLat / newAmount;
      }
      //Logger.log(plzArray[i].plz);
      newPlaces.push({
        amount : newAmount,
        lng : newLng,
        lat : newLat,
        plz : newPlz,
        connections : []
      });
    }
  }
  //Logger.log(newPlaces);
 printMap(newPlaces, "finalMap"); 
}


function printMap(newPlaces, name) {
  
   // Create a map centered on Times Square.
   var map = Maps.newStaticMap()
       .setSize(1920, 1080)
       .setCenter('Schweiz')
       .setZoom(9)
       .setMapType(Maps.StaticMap.Type.ROADMAP);

  for (var i = 0; i < newPlaces.length; i++){
    var entry = newPlaces[i];
    
    var size = getSize(entry.amount);
    var query = entry.lng + ", " + entry.lat;
    
    Logger.log(entry.lat + ", " + entry.lng + ' https://www.iconsdb.com/icons/download/soylent-red/circle-' + size + '.ico');
    
    map
      .setCustomMarkerStyle('https://www.iconsdb.com/icons/download/soylent-red/circle-' + size + '.ico', false)
      .addMarker(entry.lat, entry.lng);

    
  }
  

   Logger.log(map.getMapUrl());
  var blob = map.getBlob();
  DriveApp.getFolderById("1bcIVaYFYGW79S1Kw9vEk-bkDhShDgRrP").createFile(blob);
}