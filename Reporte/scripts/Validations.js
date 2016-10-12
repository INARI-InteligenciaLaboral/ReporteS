function JavascriptFunction() {
    var mCheck = $('#IsPeridos').val();
    if (mCheck == 'Periodos')
    {
        var mPerIni = $('#PerIni').val();
        if(!mPerIni == '')
        {
            document.getElementById('lbltipAddedComment').innerHTML = 'El documento Reporte se generara en breve, favor de esperar';
        }
        else
        {
            document.getElementById('lbltipAddedComment').innerHTML = 'Favor de coloar un periodo';
        }
    }
    else
    {
        document.getElementById('lbltipAddedComment').innerHTML = 'El documento Reporte.xlsx se generara en breve, favor de esperar';
    }
}
function functionBimestral() {
    document.getElementById('mensual').style.display = 'none';
    document.getElementById('bimestral').style.display = 'block';
    document.getElementById('periodo').style.display = 'none';
    document.getElementById('boton').style.display = 'none';
    var dano = document.getElementById('SelectedAnos');
    dano.selectedIndex = 0;
    var dper = document.getElementById('SelectedBimestre');
    dper.selectedIndex = 0;
    
}
function functionMensual() {
    document.getElementById('mensual').style.display = 'block';
    document.getElementById('bimestral').style.display = 'none';
    document.getElementById('periodo').style.display = 'none';
    document.getElementById('boton').style.display = 'none';
    var dano = document.getElementById('AnoMen');
    dano.selectedIndex = 0;
    var dper = document.getElementById('PerMen');
    dper.selectedIndex = 0;
}
function functionPeriodo() {
    document.getElementById('mensual').style.display = 'none';
    document.getElementById('bimestral').style.display = 'none';
    document.getElementById('periodo').style.display = 'block';
    document.getElementById('boton').style.display = 'block';
    document.getElementById("PerIni").value = '';
    document.getElementById("PerFin").value = '';
}
function functionValores() {
    var m_reporte = $('#SelectedBase').val();
    var m_empresa = $('#SelectedEmpresas').val();
    var m_procesos = $('#SelectedProcesos').val();
    if (!m_empresa == '' && !m_procesos == '' && !m_reporte == '')
    {
        document.getElementById('checks').style.display = 'block';
        if (m_reporte == 'siap')
        {
            document.getElementById('tablaperi').style.display = 'none';
        }
        else
        {
            document.getElementById('tablaperi').style.display = 'block';
        }
    }
    else
    {
        document.getElementById('checks').style.display = 'none';
        document.getElementById('boton').style.display = 'none';
        document.getElementById('mensual').style.display = 'none';
        document.getElementById('bimestral').style.display = 'none';
        document.getElementById('periodo').style.display = 'none';
        $("input:radio").attr("checked", false);
    }
}
function functionMenVal() {
    var m_mensual = $('#AnoMen').val();
    var m_periodo = $('#PerMen').val();
    if (document.getElementById('checks').style.display == 'block')
    {
        if(!m_mensual == '' && !m_periodo == '')
        {
            document.getElementById('boton').style.display = 'block';
        }
        else
        {
            document.getElementById('boton').style.display = 'none';
        }
    }
    else
    {
        document.getElementById('boton').style.display = 'none';
    }
}

function functionReportes() {
    var m_reporte = $('#SelectedReporte').val();
    if (!m_reporte == '') {
        document.getElementById('filtros').style.display = 'block';
    }
    else {
        document.getElementById('filtros').style.display = 'none';
        var dano = document.getElementById('SelectedBase');
        dano.selectedIndex = 0;
    }
}
function functionBimVal() {
    var m_bimestre = $('#SelectedAnos').val();
    var m_periodo = $('#SelectedBimestre').val();
    if (document.getElementById('checks').style.display == 'block') {
        if (!m_bimestre == '' && !m_periodo == '') {
            document.getElementById('boton').style.display = 'block';
        }
        else {
            document.getElementById('boton').style.display = 'none';
        }
    }
    else {
        document.getElementById('boton').style.display = 'none';
    }
}