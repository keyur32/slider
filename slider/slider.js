var _OM = null;
var _settings = null;           //settings object
var min = 0;
var max = 100;
var step = 1;
var cellLinkValue = 20;
var sliderOrientation = "horizontal";  //orientation

var isRTL = false;  //rtl
var showTickMarks = true;
var showTooltip = true;
var sliderDirection = 'ltr';
var sliderColor = "rgb(234, 37, 119)";
var colorMode = "default";
var readonly = false;
var theme = "default";
var title  = "";
var conditionalColors = ["red", "yellow", "green"];
var valueRed = 0;
var valueYellow = 0;


  	$(function() {

		InitSlider();
		setSliderValue(91.2);

		$(window).resize(function()
		{ // On resize
			SetSliderHeight();
	   	});

	   	var icons = {
		      header: "ui-icon-circle-arrow-e",
		      activeHeader: "ui-icon-circle-arrow-s"
		    };
		    $( "#accordion" ).accordion({
		      icons: icons,
		      collapsible: true});
	}); //end onload

	function GetSettingsB()
	{
		title="foobar";
		min = 5;
		max = 100;
		step = 5;
		colorMode = "default";
		theme = "windows";
		sliderColor = "purple";

	}

	function SetSliderHeight()
	{
		//add resize event for window
		if(sliderOrientation == "vertical")
		{
				var h = $(window).height()- 110;
				$('#wrapper').css('height', h);
				$('#slider').css('height', '100%');
		}
		else
		{
				$('#wrapper').css('height', '');
				$('#slider').css('height', '');
		}
	}

	function InitSlider()
	{
	    var sliderVal = min;
	    try{
	        sliderVal = $("#slider").val();
	    }
	    catch (e) {
	        sliderVal = min;
	    }


	    var connectDirection = "lower";
	    if (sliderOrientation == "vertical") {
	        sliderDirection = "rtl";

	        //set height manually
	        var h = $(window).height()- 110;
	        $('#wrapper').css('height', h);
	        $('#slider').css('height', '100%');
	        connectDirection = "lower";

	        //reset the slider so we can repaint
	        $('#slider').html("");
	    }
	    else {
	        $('#wrapper').css('height', '');
	        $('#slider').css('height', '');
	        sliderDirection = "ltr";
	    }

		//init
		$("#slider").noUiSlider({
		    start: [cellLinkValue],
            direction: sliderDirection,
		    orientation: sliderOrientation,
					connect: connectDirection,
					range: {
						'min': min,
						'max': max,
					},
					'step': step,
			}, true);

	    //reset background color
		$("#slider").css("background-color", "");

		//remove any extra classes that may have been added
		$('.noUi-handle').removeClass("theme-windows-handle-horizontal theme-windows-handle-vertical theme-ios-handle-horizontal theme-ios-handle-vertical");
		$('.noUi-target').removeClass("theme-windows-bar-horizontal theme-windows-bar-vertical theme-ios-bar-horizontal theme-ios-bar-vertical");

		if(theme == "iOS7")
		{
			//addClass to the element
			$('.noUi-handle').addClass("theme-ios-handle-"+sliderOrientation);
			$('.noUi-target').addClass("theme-ios-bar-"+sliderOrientation);
			$('.noUi-handle').removeClass("theme-windows-handle-"+sliderOrientation);
			$('.noUi-target').removeClass("theme-windows-bar-"+sliderOrientation);
		}
		else if(theme == "windows")
		{
			$('.noUi-base').addClass("theme-windows");
			$('.noUi-origin').addClass("theme-windows");
			$('.noUi-handle').addClass("theme-windows-handle-"+sliderOrientation);
			$('.noUi-target').addClass("theme-windows-bar-"+sliderOrientation);
		}
		else
		{
			$('.noUi-handle').removeClass("theme-ios-handle-"+sliderOrientation+" theme-windows-handle-"+sliderOrientation);
			$('.noUi-target').removeClass("theme-ios-bar-"+sliderOrientation+" theme-windows-bar-"+sliderOrientation);
		}

		$('.noUi-connect').css('background-color', sliderColor);


		//disable
		if (readonly) {
			$('#slider').attr('disabled', 'disabled');
		}
		else {
			$('#slider').removeAttr('disabled');
		}

		//add events
		$("#slider").on({
					slide: function(){

						cellLinkValue = $('#slider').val();
						setTooltipValue(cellLinkValue);
						if(colorMode == "conditional")
						SetConditionalColor(cellLinkValue);
					},
					set: function(){

					},
					change: function(){

						//write back to Excel only after it's completed
						cellLinkValue = $('#slider').val();
						WriteData(cellLinkValue);
						if(colorMode == "conditional")
						SetConditionalColor(cellLinkValue);
			}
		});


		if (showTickMarks) {
		    $('#slider').noUiSlider_pips({
		        mode: 'count',
		        values: 6,
		        density: 4
		    }, true);
		}

		if (showTooltip) {
		    /*add tooltip*/
		    var tooltipHtml = '<div class="tooltip"><div class="tooltip-inner">' + cellLinkValue + '</div><div class="tooltip-arrow"></div></div>';
		    var $handle = $("#slider").find(".noUi-handle");
		    $handle.append(tooltipHtml);
		    setTooltipPosition();
		}
	}


	function SetConditionalColor(cellLinkValue)
	{
		var top3rd = (max-min)/3;
		if(!valueRed)
			valueRed = top3rd;
		if(!valueYellow)
			velueYellow = top3rd*2;

		if(cellLinkValue <= valueRed)
		{
			sliderColor = conditionalColors[0];
		}
		else if(cellLinkValue <= valueYellow)
		{
			sliderColor = conditionalColors[1];
		}
		else
		{
			sliderColor = conditionalColors[2];
		}
		  //reset background color
		  $("#slider").css("background-color", "");
		  $('.noUi-connect').css('background-color', sliderColor);
	}

	//bind to cell in Excel so we can write to and recieve data changed events

	function Bind() {

	    //create binding to current cell in Active Excel workbook
	    _OM.bindings.addFromPromptAsync(Office.BindingType.Text,
	             {
	                 id: "binding1",
	                 promptText: "Select a cell to link the slider to."
	             },
	             function (asyncResult) {
	                 if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {

	                     var binding = asyncResult.value;
	                     binding.addHandlerAsync(Office.EventType.BindingDataChanged, DataChanged);
	                     GetData(binding);
	                 }
	             });
	}

	//Handle DataChanged Event
	function DataChanged(asyncResult) {

	    GetData();
	}

	function GetData()
	{
		var binding = Office.select("bindings#binding1");
	    binding.getDataAsync(
		{
		    coercionType: Office.CoercionType.Text,
		    valueFormat: Office.ValueFormat.Unformatted,
		    filterType: Office.FilterType.OnlyVisible
		},
		ChartData);

	}

	function ChartData(asyncResult) {

	    //TODO: add error check
	    if (asyncResult.status == "succeeded") {
	        var data = asyncResult.value;
	        setSliderValue(data);
	    }
	    else {
	       log("Uh-oh, something happened that we weren't expecting. Error Message: " + asyncResult.error.message);
	    }
	}

	function WriteData(value)
	{
		log("Write data");
		//write cell value back to binding
		var binding = Office.select("bindings#binding1");
		binding.setDataAsync(value, function (asyncResult) {
			});
	}

	function setSliderValue(value)
	{
		$("#slider").val(value);
		setTooltipValue(value);
	}

	function setTooltipPosition()
	{
		//calculate left position based tooltip length
		var $handle = $("#slider").find(".noUi-handle");
		var tooltip = document.getElementsByClassName('tooltip')[0];
		var offset = Math.abs($handle[0].clientWidth - tooltip.clientWidth)/ 2;
		tooltip.style.left = -1*offset + "px";

	}

	function setTooltipValue(value)
	{
		var tooltip = document.getElementsByClassName('tooltip-inner')[0];
		tooltip.innerHTML = value;
		setTooltipPosition();
	}

	function ShowColorPicker(cp, index)
	{
		var $connectBar = $('.noUi-connect');

			$('.color-box').colpick(
			{
				color: cp.style.background,
				layout:'hex',
				submit:0,
				onChange:function(hsb,hex,rgb,el,bySetColor) {

					$(el).css('background-color', '#'+hex);
					//$(el).colpickHide();

				}
			}).keyup(function(){
				$(this).colpickSetColor(this.value);
			});


	}

	function UpdateColorSelect()
	{
		colorMode = $('#formatStyle').val();
		if(colorMode == "default") //one color
		{
			$('.FormatStyleFixedRow').css('display', '');
			$('.FormatStyleScaleRow').css('display', 'none');
		}
		else
		{
			$('.FormatStyleFixedRow').css('display', 'none');
			$('.FormatStyleScaleRow').css('display', 'block');


			SetConditionalLabels();
		}
	}

	function SetConditionalLabels()
	{
		//set labels
		min = parseFloat($('#min').val());
		max = parseFloat($('#max').val());
		var interval = (max-min)/3;
		$('#valueRed').val(interval);
		$('#valueYellow').val(interval*2);
		$('#valueGreen').html(max);

	}

	function GetSettings() {

		if(!_OM)
			return;
		//get settings
		_settings = _OM.settings;
		min = _settings.get("min");
		max = _settings.get("max");
		step = _settings.get("step");
		title = _settings.get("title");

		log("**title: "+title+ " **step: "+step+" max:"+max);
		if(!min)
			min = 0;

		if(!max)
			max = 100;
		if(!step)
			step = .01;
		if(!title){
			log('reset title');
			title = "";
		}

		$('#titleToShow').val(title);
		$('#title').html(title);


		//get and populate from settings object
		theme = _settings.get("theme");
		if(!theme || theme.length == 0)
			theme = "default";

		colorMode = _settings.get("colorMode");
		if(!colorMode)
			colorMode = "single";

		if(colorMode == "conditional")
		{
			conditionalColors[0] = _settings.get("colorRed");
			conditionalColors[1] = _settings.get("colorYellow");
			conditionalColors[2] = _settings.get("colorGreen");
			valueRed = _settings.get("valueRed");
			valueYellow = _settings.get("valueYellow");

			if(!conditionalColors[0]) conditionalColors[0] = "red";
			if(!conditionalColors[1]) conditionalColors[1] = "yellow";
			if(!conditionalColors[2]) conditionalColors[2] = "green";

			//valueRed
			//valueYellow = valueRed * 2;
		}
		else
		{
			sliderColor = _settings.get("sliderColor");
		}

		sliderOrientation = _settings.get("sliderOrientation");
		showTickMarks =	_settings.get("showTickMarks");
		showTooltip = _settings.get("showTooltip");
		readonly =	_settings.get("readonly");




		log(" sliderColor:"+sliderColor);
		log(" readonly:"+readonly);
		log(" showTickMarks: "+ showTickMarks);
		log(" showTooltip: " + showTooltip);
		log(" sliderOrientation: "+ sliderOrientation);
	}

	function SetSettings() {

		_settings = _OM.settings;

		//data
		_settings.set("min", min);
		_settings.set("max", max);
		_settings.set("step", step);
		_settings.set("title", title);

		//look and feel
		_settings.set("theme", theme);
		_settings.set("colorMode", colorMode);
		if(colorMode == "conditional")
		{
			_settings.set("colorRed", conditionalColors[0]);
			_settings.set("colorYellow", conditionalColors[1]);
			_settings.set("colorGreen", conditionalColors[2]);
			_settings.set("valueRed", valueRed);
			_settings.set("valueYellow", valueYellow);
		}
		else
		{
			_settings.set("sliderColor", sliderColor); // = $('#colorBox').css('background-color');
		}


		//orientation
		_settings.set("sliderOrientation", sliderOrientation);
		_settings.set("showTickMarks", showTickMarks);
		_settings.set("showTooltip", showTooltip);
		_settings.set("readonly", readonly);
		_settings.saveAsync();

	}

	function UpdateSettings()
	{
        //get settings
		sliderColor = $('#color-box').css('background-color');

		min = parseFloat($('#min').val());
		max = parseFloat($('#max').val());
		step = parseFloat($('#step').val());

		showTickMarks = $('#tick').is(':checked');
		showTooltip = $('#tip').is(':checked');
		readonly = $('#readonly').is(':checked');

		sliderOrientation = $('#orientation').val();
		title = $('#titleToShow').val();
		$('#title').html(title);
		theme = $('#theme').val();
		colorMode = $('#formatStyle').val();

		if(colorMode == "conditional") {

			conditionalColors[0] = $('#colorBoxRed').css('background-color');
			conditionalColors[1] = $('#colorBoxYellow').css('background-color');
			conditionalColors[2] = $('#colorBoxGreen').css('background-color');

			valueRed = parseFloat($('#valueRed').val());
			valueYellow = parseFloat($('#valueYellow').val());

			SetConditionalColor(cellLinkValue);
		}
		else
		{
			sliderColor = $('#colorBox').css('background-color');
		}
	    //save them to Settings

        //reset UI
		ToggleSettings();
		if(_OM) SetSettings();
		InitSlider();

	}

	function log(text)
	{
		var debug = false;
		if(debug)
			$('#log').html($('#log').html() + "<br />" + text);
	}

	function ShowSettings()
	{


		$('#min').val(min);
		$('#max').val(max);
		$('#step').val(step);
		$('#theme').val(theme);
		$('#formatStyle').val(colorMode);
		$('#tip').prop('checked', showTooltip);
		$('#tick').prop('checked', showTickMarks);
		$('#readonly').prop('checked', readonly);
		$('#colorBox').css('background-color', sliderColor);
	}

	function ToggleSettings()
	{
		if($('#settings').css('display') == 'none') {
			$('#settings').css('display', 'block');
			$('#wrapper').css('display', 'none');
			$('#settings-button').css('display', 'none');
			$('#title').css('display', 'none');

		}
		else {

			$('#settings').css('display', 'none');
			$('#wrapper').css('display', 'block');
			$('#settings-button').css('display', 'block');
			$('#title').css('display', 'block');
			}
	}

	function bodyOnMouseOver()
	{
		$('#settings-button').css('visibility', 'visible');

	}

	function bodyOnMouseOut()
	{
		$('#settings-button').css('visiblity', 'hidden');
	}

	// Add any initialization logic to this function.
	Office.initialize = function (reason) {

	    // Checks for the DOM to load.
	    $(document).ready(function () {

	        //setup globals as needed
	        _OM = Office.context.document;
	        log("Getting Office context" + _OM);

			GetSettings();
	        InitSlider();


	        //check if binding exists...
	        Office.context.document.bindings.getByIdAsync('binding1', function (asyncResult) {

	            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
	                var binding = asyncResult.value;
	                binding.addHandlerAsync(Office.EventType.BindingDataChanged, DataChanged);

	                GetData();
	            }
	        });
	    });
}