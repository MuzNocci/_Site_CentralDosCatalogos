function camposRequeridosOu(form, nomes, rotulos) {
  var valido = false;
  for (var i = 0; i < nomes.length; ++i) {
    var campo = form[nomes[i]];
    if (campo != null) {
      var value = campo.value;
      if (value != null && value.length > 0) {
        valido = true;
        break;
      }
    }
  }
  if (!valido) {
    var msg = 'Um dos campos deve ser preenchido: ';
    for (var i = 0; i < nomes.length; ++i) {
      if (i > 0) {
        msg = msg + ' ou ';
      }
      msg = msg + rotulos[i];
    }
    throw msg;
  }
}

function camposRequeridosE(form, nomes, rotulos) {
  var validar = false;
  for (var i = 0; i < nomes.length; ++i) {
    var campo = form[nomes[i]];
    if (campo != null) {
      var value = campo.value;
      if (value != null && value.length > 0) {
        validar = true;
      }
    }
  }
  if (validar) {
    for (var i = 0; i < nomes.length; ++i) {
      campoRequerido(form, nomes[i], rotulos[i]);
    }
  }
}

function campoRequerido(form, nome, rotulo) {
  var campo = form[nome];
  if (campo != null) {
    var value = campo.value;
    if (value == null || value.length == 0) {
      campo.focus();
      throw rotulo + ' é de preenchimento obrigatório!';
    }
  }
}

function total(indice, qtdLinhas) {
  var form, itQtd, itUnit, itTotal, qtd, unit, total, totalStr, index;
  form = document.forms[0];
  itQtd = form["str_qtd" + indice];
  qtd = parseFloat(itQtd.value);
  itUnit = form["str_unitario" + indice];
  unit = parseFloat(itUnit.value.replace('.', '').replace(',', '.'));
  itTotal = form["str_total" + indice];
  total = qtd * unit;
  totalStr = ajustarDecimais(total);
  itTotal.value = formatNumber(totalStr, '##.##0,00');
  var totalGeral, itTotalGeral;
  totalGeral = 0;
  itTotalGeral = form["str_totalGeral"];
  if (itTotalGeral != null) {
    for (var i = 1; i <= qtdLinhas; ++i) {
      var tot = form["str_total" + i]
      if (tot != null && tot.value != "") {
        try {
          totalGeral = totalGeral + parseFloat(tot.value.replace('.', '').replace(',', '.'));
        } catch (Err) {
        }
      }
    }
    totalStr = ajustarDecimais(totalGeral);
    itTotalGeral.value = "R$ " + formatNumber(totalStr, '##.##0,00');
  }
}

function ajustarDecimais(total) {
  var totalStr;
  totalStr = String(total);
  index = totalStr.indexOf(".");
  if (index > -1) {
    var qtdDecimais = totalStr.length - index - 1;
    if (qtdDecimais > 2) {
      totalStr = totalStr.substring(0, index + 3);
    } else if (qtdDecimais < 2) {
      totalStr = totalStr + "0";
    }
  } else {
      totalStr = totalStr + "00";
  }
  return totalStr;
}