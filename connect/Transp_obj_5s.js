
async function container_5s (Labels,Title,Subtitle,DataSets) {
  const json_cont = {
      error : "no",
      labels : Labels,
      title : Title,
      subtitle : Subtitle,
      datasets : [
        {
          dt_nam:"Wynik 5s",
          fillColor : "rgba(151,187,205,0.5)",
          strokeColor : "rgba(151,187,205,1)",
          pointColor : "rgba(220,220,220,1)",
          pointStrokeColor : "#fff",
          jedn : "%",
          j_nam : "Procent",
          data : DataSets[0]
        }],
      menu_item : [
        { capt: 'Wyniki audytów 5s',
          dtaset: 0,
        }
      ],      
    };
  return json_cont;
}

module.exports = {
  Create: (Labels,Title,Subtitle,DataSets) => container_5s(Labels,Title,Subtitle,DataSets)
};
