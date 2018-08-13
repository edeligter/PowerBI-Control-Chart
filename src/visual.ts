/*
 *  Power BI Visualizations
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved. 
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *   
 *  The above copyright notice and this permission notice shall be included in 
 *  all copies or substantial portions of the Software.
 *   
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */

module powerbi.extensibility.visual {
    /**
         * Interface for ControlChart viewmodel.
         *
         * @interface
         * @property {any[]} points                 - Set of data points the visual will render.
         * @property {any} minX                     - minimum value of X axis - can be date or number
         * @property {any} maxX                     - maximum value of X axis - can be date or number
         * @property {any} minY                     - minimum value of Y axis - can be number
         * @property {any} maxY                     - maximum value of Y axis - can be number
         * @property {interface} data               - LineData - data point line style
         * @property {string} backgroundColor       - Plot area color
         * @property {interface} marker             - Marker size
         * @property {interface} gridlines          - Gridline style
         * @property {interface} subgroupDividerLine   - StatisticsData - contains info on line and labels
         * @property {interface} limitLine          - StatisticsData - contains info on line and labels
         * @property {interface} meanLine           - StatisticsData - contains info on line and labels
         * @property {interface} xAxis              - AxisData - contains info on label style
         * @property {interface} yAxis              - AxisData - contains info on label style
         * @property {boolean} isDateRange          - is the X axis a date or numeric range?
         * @property {boolean} runRule1             - run rule 1
         * @property {boolean} runRule2             - run rule 2
         * @property {boolean} runRule3             - run rule 3
         * @property {string} ruleColor             - color of points matching rule
         * @property {number} movingRange           - moving range window width
         * @property {boolean} mRError              - moving range is larger than number of data points in a Subgroup
         * @property {boolean} showmRWarning        - toggle error message if mRError
         */
    interface ControlChartViewModel {
        points: any[],
        minX: any;
        maxX: any;
        minY: number;
        maxY: number;
        data: LineData;
        backgroundColor: string;
        marker: MarkerStyle;
        gridlines: GridLinesStyle;
        subgroupDividerLine: StatisticsData;
        limitLine: StatisticsData;
        meanLine: StatisticsData;
        standardDeviations: number;
        xAxis: AxisData;
        yAxis: AxisData;
        isDateRange: boolean;
        xAxisFormat: any;
        runRule1: boolean;
        runRule2: boolean;
        runRule3: boolean;
        ruleColor: string;
        movingRange: number;
        mRError: boolean;
        showmRWarning: boolean;
    };

    interface LineData {
        DataColor: string;
        LineColor: string;
        LineStyle: string;
    }

    interface MarkerStyle{
        MarkerSize: number;       
    }

    interface AxisData {
        AxisTitle: string;
        TitleSize: number;
        TitleFont: string;
        TitleColor: string;
        AxisLabelSize: number;
        AxisLabelFont: string;
        AxisLabelColor: string;
        AxisLabelFormat: any;
        Rotation?: number;
    }

    interface StatisticsData {
        textSize: number;
        textFont: any;
        textColor: string;
        lineColor: string;
        lineStyle: string;
        show: boolean;
    }

    interface GridLinesStyle{
        lineColor: string;
        lineStyle: string;
        show: boolean;
    }

    /**
         * Interface for ControlChart Stage.
         *
         * @interface
         * @property {number} UCL           - Upper control limit.
         * @property {number} LCL           - Lower control limit.
         * @property {number} Mean          - Mean
         * @property {number} startX        - first point in stage
         * @property {number} endX          - last point in stage
         * @property {number} sum           - sum of data in stage
         * @property {number} count         - number of points in stage
         * @property {string} stage         - Stage name
         * @property {any} stageDividerX    - stage divider x axis value - date or number
         * @property {number} firstId       - first index value in a stage - use this in calculating stats
         * @property {number} lastId        - first index value in a stage - use this in calculating stats
         * @property {boolean} mrError      - if moving range is >= number of data points in each stage
         */
    interface Subgroup {
        lCL: number;
        uCL: number;
        mean: number;
        startX: any;
        endX: any;
        sum: number;
        count: number;
        stage: string;
        stageDividerX: any;
        firstId: number;
        lastId: number;
        mRError: boolean;
    };

    /**
         * Function that converts queried data into a view model that will be used by the visual.
         *
         * @function
         * @param {VisualUpdateOptions} options - Contains references to the size of the container
         *                                        and the dataView which contains all the data
         *                                        the visual had queried.
         * @param {IVisualHost} host            - Contains references to the host which contains services
         */
    function visualTransform(options: VisualUpdateOptions, host: IVisualHost): ControlChartViewModel {
        let dataViews = options.dataViews;
        let viewModel: ControlChartViewModel = {
            points: [],
            minX: null,
            maxX: null,
            minY: 0,
            maxY: 0,
            data: null,
            backgroundColor: null,
            marker: null,
            gridlines: null,
            xAxis: null,
            yAxis: null,
            subgroupDividerLine: null,
            meanLine: null,
            limitLine: null,
            standardDeviations: 3,
            isDateRange: true,
            xAxisFormat: null,
            runRule1: false,
            runRule2: false,
            runRule3: false,
            ruleColor: null,
            movingRange: 2,
            mRError: false,
            showmRWarning: true
        };

        let defGridlineColor: string = '#c2c6c6';
        let defAxisLabelColor: string = '#000000';
        let defMeanLineColor: string = '#35BF4D';
        let defSubgroupLineColor: string = '#00C3FF';
        let defLimitLineColor: string = '#FFA500';

        if (!dataViews
            || !dataViews[0]
            || !dataViews[0].categorical
            || !dataViews[0].categorical.categories
            || !dataViews[0].categorical.categories[0].source
            || !dataViews[0].categorical.values
            || !dataViews[0].categorical.values[0].source
        )
            return viewModel;

        let categorical = dataViews[0].categorical;
        let category = categorical.categories[0];
        let dataValue = categorical.values[0];
        let Points: any[] = [];  
        var xValues: PrimitiveValue[] = category.values;
        var yValues: PrimitiveValue[] = dataValue.values;
        var isYAxisNumericData: boolean = false;// (dataValue.values[0] && Object.prototype.toString.call(dataValue.values[0]) === '[object Number]');
        var xAxisType: string = "NotUsable";
        
        try{
            
            if (category.source.type.dateTime.valueOf() == true)
                xAxisType = "date";
            else
                if(category.source.type.numeric.valueOf() == true)
                    xAxisType = "numeric";
            
            if (dataValue.source.type.numeric.valueOf() == true)
                isYAxisNumericData = true;
    
            if (xAxisType != "NotUsable" && isYAxisNumericData) {
                for (let i = 0; i < xValues.length; i++) {
                    Points.push([xValues[i],<number>yValues[i]]);
                }

                let dvobjs = dataViews[0].metadata.objects;

                var xAxisLabelFormat: any;
                if (xAxisType == "date")
                    xAxisLabelFormat = getValue<string>(dvobjs, 'xAxis', 'xAxisDateFormat', '%d-%b-%y');
                else
                    xAxisLabelFormat = getValue<string>(dvobjs, 'xAxis', 'xAxisLabelFormat', '.3s')

                let chartData: LineData = {
                    DataColor: getFill(dataViews[0], 'chart', 'dataColor', '#FF0000'),
                    LineColor: getFill(dataViews[0], 'chart', 'lineColor', '#0000FF'),
                    LineStyle: getValue<string>(dvobjs, 'chart', 'lineStyle', '')
                };
                let meanLine: StatisticsData = {
                    textColor: getFill(dataViews[0], 'statistics', 'meanLabelColor', defMeanLineColor),
                    textSize: getValue<number>(dvobjs, 'statistics', 'meanLabelSize', 10),
                    textFont: getValue<string>(dvobjs, 'statistics', 'meanLabelfontFamily', 'Arial'),
                    lineColor: getFill(dataViews[0], 'statistics', 'meanLineColor', defMeanLineColor),
                    lineStyle: getValue<string>(dvobjs, 'statistics', 'meanLineStyle', '6,4'),
                    show: getValue<boolean>(dvobjs, 'statistics', 'showMean', true)
                };
                let stageDividerLine: StatisticsData = {
                    textColor: getFill(dataViews[0], 'subgroups', 'subgroupLabelColor', defSubgroupLineColor),
                    textSize: getValue<number>(dvobjs, 'subgroups', 'subgroupDividerLabelSize', 12),
                    textFont: getValue<string>(dvobjs, 'subgroups', 'subgroupDividerfontFamily', 'Arial'),
                    lineColor: getFill(dataViews[0], 'subgroups', 'subgroupDividerColor', defSubgroupLineColor),
                    lineStyle: getValue<string>(dvobjs, 'subgroups', 'subgroupDividerLineStyle', '10,4'),
                    show: getValue<boolean>(dvobjs, 'subgroups', 'showDividers', true)
                };
                let limitLine: StatisticsData = {
                    textColor: getFill(dataViews[0], 'subgroups', 'limitLabelColor', defLimitLineColor),
                    textSize: getValue<number>(dvobjs, 'subgroups', 'limitLabelSize', 10),
                    textFont: getValue<string>(dvobjs, 'subgroups', 'limitLabelfontFamily', 'Arial'),
                    lineColor: getFill(dataViews[0], 'subgroups', 'limitLineColor', defLimitLineColor),
                    lineStyle: getValue<string>(dvobjs, 'subgroups', 'limitLineStyle', '6,4'),
                    show: getValue<boolean>(dvobjs, 'subgroups', 'showLimits', true)
                };
                let xAxisData: AxisData = {
                    AxisTitle: getValue<string>(dvobjs, 'xAxis', 'xAxisTitle', 'Default Value'),
                    TitleColor: getFill(dataViews[0], 'xAxis', 'xAxisTitleColor', defAxisLabelColor),
                    TitleSize: getValue<number>(dvobjs, 'xAxis', 'xAxisTitleSize', 12),
                    TitleFont: getValue<string>(dvobjs, 'xAxis', 'xAxisTitlefontFamily', 'Arial'),
                    AxisLabelSize: getValue<number>(dvobjs, 'xAxis', 'xAxisLabelSize', 12),
                    AxisLabelColor: getFill(dataViews[0], 'xAxis', 'xAxisLabelColor', defAxisLabelColor),
                    AxisLabelFont: getValue<string>(dvobjs, 'xAxis', 'xAxisLabelfontFamily', 'Arial'),
                    Rotation: getValue<number>(dvobjs, 'xAxis', 'xAxisLabelRotation', 0),
                    AxisLabelFormat: xAxisLabelFormat
                };
                let yAxisData: AxisData = {
                    AxisTitle: getValue<string>(dvobjs, 'yAxis', 'yAxisTitle', 'Default Value'),
                    TitleColor: getFill(dataViews[0], 'yAxis', 'yAxisTitleColor', defAxisLabelColor),
                    TitleSize: getValue<number>(dvobjs, 'yAxis', 'yAxisTitleSize', 12),
                    TitleFont: getValue<string>(dvobjs, 'yAxis', 'yAxisTitlefontFamily', 'Arial'),
                    AxisLabelSize: getValue<number>(dvobjs, 'yAxis', 'yAxisLabelSize', 12),
                    AxisLabelFont: getValue<string>(dvobjs, 'yAxis', 'yAxisLabelfontFamily', 'Arial'),
                    AxisLabelColor: getFill(dataViews[0], 'yAxis', 'yAxisLabelColor', defAxisLabelColor),
                    AxisLabelFormat: getValue<string>(dvobjs, 'yAxis', 'yAxisLabelFormat', '.3s')
                };
                let Marker: MarkerStyle = {
                    MarkerSize: getValue<number>(dvobjs, 'chart', 'markerSize', 3),
                }
                let GridLines: GridLinesStyle = {
                    lineColor: getFill(dataViews[0], 'chart', 'gridlinesColor', defGridlineColor),
                    lineStyle: getValue<string>(dvobjs, 'chart', 'gridlinesStyle', '1,4'),
                    show: getValue<boolean>(dvobjs, 'chart', 'showGridLines', true)
                }

                var mRange: number = getValue<number>(dvobjs, 'statistics', 'movingRange', 2);
                if (mRange < 2 || mRange > 50)
                    mRange = 2
                else
                    mRange = Math.round(mRange);

                return {                    
                    points: Points,
                    minX: null,
                    maxX: null,
                    minY: 0,
                    maxY: 0,
                    data: chartData,
                    backgroundColor: getFill(dataViews[0], 'chart', 'backgroundColor', '#FFFFFF'),
                    marker: Marker,
                    gridlines: GridLines,
                    xAxis: xAxisData,
                    yAxis: yAxisData,
                    subgroupDividerLine: stageDividerLine,
                    limitLine: limitLine,
                    meanLine: meanLine,
                    standardDeviations: getValue<number>(dvobjs, 'statistics', 'standardDeviations', 3),             
                    isDateRange: (xAxisType == "date"),
                    xAxisFormat: xAxisLabelFormat,
                    runRule1: getValue<boolean>(dvobjs, 'rules', 'runRule1', false),
                    runRule2: getValue<boolean>(dvobjs, 'rules', 'runRule2', false),
                    runRule3: getValue<boolean>(dvobjs, 'rules', 'runRule3', false),
                    ruleColor: getFill(dataViews[0], 'rules', 'ruleColor', '#FFFF00'),
                    movingRange: mRange,
                    mRError: false,
                    showmRWarning: getValue<boolean>(dvobjs, 'statistics', 'showmRWarning', true)
                }
            }
            else {
                return viewModel;   // --- can null be returned????
            }
        }
        catch(e){
            return viewModel;
        }
    }

    export class ControlChart implements IVisual {
        private host: IVisualHost;
        private dataView: DataView;
        private subgroups: Subgroup[];
        private controlChartViewModel: ControlChartViewModel;
        private svgRoot: d3.Selection<SVGElementInstance>;
        private svgGroupMain: d3.Selection<SVGElementInstance>;
        private padding: number = 12;
        private plot;
        private xScale;
        private yScale;
        private meanLine = [];
        private uclLines = [];
        private lclLines = [];
        private subgroupDividers = [];
        private dots;
        private tooltipServiceWrapper: ITooltipServiceWrapper;

        constructor(options: VisualConstructorOptions) {
            this.host = options.host;
            this.svgRoot = d3.select(options.element).append('svg').classed('controlChart', true);
            this.svgGroupMain = this.svgRoot.append("g").classed('Container', true);
            this.tooltipServiceWrapper = createTooltipServiceWrapper(this.host.tooltipService, options.element);
        }

        public update(options: VisualUpdateOptions) {
            // remove all existing SVG elements 
            this.svgGroupMain.selectAll("*").remove();
            this.svgRoot.empty();
            if (!options.dataViews[0]
                || !options.dataViews[0].categorical
                || !options.dataViews[0].categorical.categories
                || !options.dataViews[0].categorical.values)
                return;

            this.dataView = options.dataViews[0];
            // convert categorical data into specialized data structure for data binding
            this.controlChartViewModel = visualTransform(options, this.host);
            this.svgRoot
                .attr("width", options.viewport.width)
                .attr("height", options.viewport.height);
            
            if (this.controlChartViewModel && this.controlChartViewModel.points[0]) { //check if this detects empty viewModel
                var plot = this.plot;
                this.GetSubgroups();                                                //determine subgroups
                this.CalcStats();                                                   //calc mean and sd
                this.CreateAxes(options.viewport.width, options.viewport.height); 
                if (this.controlChartViewModel.meanLine.show)
                    this.PlotMean();                                                //mean line
                if (this.controlChartViewModel.subgroupDividerLine.show)
                    this.DrawSubgroupDividers();                                    //subgroup changes
                if (this.controlChartViewModel.limitLine.show)
                    this.PlotControlLimits();                                       //lcl and ucl
                this.PlotData();                                                    //plot basic raw data --- need to check for valid data
                this.ApplyRules();
                this.DrawMRWarning();
            }
        }      
        
        private RotationTranslate(angle: number, label:any){
            var radAngle: number = angle * Math.PI/180;
            var sinAngle = Math.sin(radAngle);
            var lenText: number = label.toString().length;
            var textSize: number = this.controlChartViewModel.xAxis.AxisLabelSize;
            var xOffset: number = textSize * sinAngle;
            var yOffset: number = Math.abs(sinAngle) * (lenText) + xOffset/4;
            return xOffset + "," + yOffset;
        }

        
        private CreateAxes(viewPortWidth: number, viewPortHeight: number) {
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            var xAxisOffset: number = 80; //60
            var yAxisOffset: number = 54;

            var plot = {
                xAxisOffset: xAxisOffset,
                yAxisOffset: yAxisOffset,
                xOffset: this.padding + xAxisOffset,
                yOffset: this.padding,
                width: viewPortWidth - (this.padding + xAxisOffset) * 2, 
                height: viewPortHeight - (this.padding * 2) - yAxisOffset,
            };
            
            var radAngle: number = viewModel.xAxis.Rotation * Math.PI/180;
            var sinAngle = Math.sin(radAngle);
            var yOffset: number = Math.abs(sinAngle) * (viewModel.xAxis.AxisLabelSize * 2); // 
            plot.height = plot.height - yOffset;

            this.plot = plot;

            this.svgGroupMain.attr({
                height: plot.height,
                width: plot.width,
                transform: 'translate(' + plot.xOffset + ',' + plot.yOffset + ')'
            });
            
            //add shading area to plot area
            this.svgGroupMain.append("rect")
                .attr("x", 0)
                .attr("y", 0)
                .attr("height", plot.height)
                .attr("width", plot.width)
                .style("fill", viewModel.backgroundColor)
                .style("stroke-width", 0);
           
                      
            let vmXaxis = viewModel.xAxis;
            let vmYaxis = viewModel.yAxis;

            this.GetMinMaxX();
            var xScale;
            var xFormat;            
            if (viewModel.isDateRange) {
                xFormat = d3.time.format(viewModel.xAxisFormat);
                xScale = d3.time.scale()
                    .range([0, plot.width])
                    .domain([viewModel.minX, viewModel.maxX])
                    .nice();
            }
            else {
                xScale = d3.scale.linear()
                    .range([0, plot.width])
                    .domain([viewModel.minX, viewModel.maxX])
                    .nice();
                xFormat = d3.format(viewModel.xAxisFormat);                
            }

            this.xScale = xScale;           
            // draw x axis
            var xAxis = d3.svg.axis()
                .scale(xScale)
                .orient('bottom')
                //.tickFormat(xFormat);
                .tickFormat(function (d) { return xFormat(d) });

            this.svgGroupMain
                .append('g')
                .attr('class', 'x axis')
                .attr('transform', 'translate(0,' + plot.height + ')')
                .call(xAxis)
                .selectAll("text")
                //.attr("transform", "translate(" + vmXaxis.Rotation/90 * 12 + " , 0 )rotate(" + vmXaxis.Rotation + ")")
                .attr("transform", "translate(" + this.RotationTranslate(vmXaxis.Rotation, function(d){return d.toString()}) + ")rotate(" + vmXaxis.Rotation + ")")// (180/ Math.PI) * Math.cos( Math.PI/180 * vmXaxis.Rotation)/2,0)})
                .style("text-anchor", "middle")
                .style('fill', vmXaxis.AxisLabelColor)
                .style("font-size", vmXaxis.AxisLabelSize + 'px')
                .style("font-family", vmXaxis.AxisLabelFont);
           
            //x grid lines
            this.svgGroupMain.append("g")			
                .attr("class", "grid")
                .attr("transform", "translate(0," + plot.height + ")")
                .call(xAxis
                    .tickSize(-plot.height)
                    .tickFormat(""));
                      
            //top and right frame lines
            this.svgGroupMain.append("g").append("line")          // attach a line
                .style("stroke", "lightgrey")  // colour the line
                .style("stroke-width", 1)
                .style("fill", "none")
                .style("shape-rendering", "crispEdges")
                .attr("x1", 0)     // x position of the first end of the line
                .attr("y1", 0)      // y position of the first end of the line
                .attr("x2", plot.width)     // x position of the second end of the line
                .attr("y2", 0);
         
            this.svgGroupMain.append("g").append("line")          // attach a line
                .style("stroke", "lightgrey")  // colour the line
                .style("stroke-width", 1)
                .style("fill", "none")
                .style("shape-rendering", "crispEdges")
                .attr("x1", plot.width)     // x position of the first end of the line
                .attr("y1", 0)      // y position of the first end of the line
                .attr("x2", plot.width)     // x position of the second end of the line
                .attr("y2", plot.height);


            //handle uCL and lCL
            var dataMax = d3.max(viewModel.points, function (d) { return d[1] });
            var dataMin = d3.min(viewModel.points, function (d) { return d[1] });

            var lclLines = this.lclLines;
            var uclLines = this.uclLines;
            var yMax: number = d3.max(uclLines, function (d) { return d['y1'] });
            var yMin: number = d3.min(lclLines, function (d) { return d['y1'] });
            yMin = Math.min(dataMin, yMin);
            if (yMin < 0)
                yMin = yMin * 1.05;
            else
                yMin = yMin * 0.95;

            yMax = Math.max(dataMax, yMax);
            if (yMax > 0)
                yMax = yMax * 1.05;
            else
                yMax = yMax * 0.95;

            // draw y axis
            var yScale = d3.scale.linear()
                .range([plot.height, 0])
                .domain([yMin, yMax])
                .nice();
            this.yScale = yScale;
            this.controlChartViewModel.minY = yMin;
            this.controlChartViewModel.maxY = yMax;

            var yformatValue = d3.format(vmYaxis.AxisLabelFormat);
            var yAxis = d3.svg.axis()
                .scale(yScale)
                .orient('left')
                .tickFormat(function (d) { return yformatValue(d) });

            this.svgGroupMain
                .append('g')
                .attr('class', 'y axis')
                .call(yAxis)
                .selectAll("text")                
                .style('fill', vmYaxis.AxisLabelColor)
                .style("font-size", vmYaxis.AxisLabelSize + 'px')
                .style("font-family", vmYaxis.AxisLabelFont);
            
            //y grid lines
            this.svgGroupMain.append("g")			
                .attr("class", "grid")
                .call(yAxis
                    .tickSize( -plot.width) 
                    .tickFormat("")                                          
                )               
                .selectAll(".tick").each(function(d,i){if (d==0 ) this.remove();});
            
            //axes titles
            this.svgGroupMain.append("text")
                .attr("transform", "rotate(-90)")
                .attr("y", 0 - xAxisOffset - 10)
                .attr("x", 0 - (plot.height / 2))
                .attr("dy", "1em")
                .style("text-anchor", "middle")
                .style("font-size", vmYaxis.TitleSize + 'px')
                .style("fill", vmYaxis.TitleColor)
                .style("font-family", vmYaxis.TitleFont)
                .text(vmYaxis.AxisTitle);
            this.svgGroupMain.append("text")
                .attr("y", plot.height + yAxisOffset)
                .attr("x", (plot.width / 2))
                .style("text-anchor", "middle")
                .style("font-size", vmXaxis.TitleSize + 'px')
                .style("fill", vmXaxis.TitleColor)
                .style("font-family", vmXaxis.TitleFont)
                .text(vmXaxis.AxisTitle);
            
            this.svgGroupMain.selectAll(".grid").attr("stroke-dasharray",  viewModel.gridlines.lineStyle);
            this.svgGroupMain.selectAll(".grid .tick").style("stroke", viewModel.gridlines.lineColor)
            if (!viewModel.gridlines.show) 
                this.svgGroupMain.selectAll(".grid").remove();
 
        }

        private GetMinMaxX() {
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            let dp = viewModel.points;
            var maxValue: any;
            var minValue: any;

            if (viewModel.isDateRange) {
                minValue = new Date();
                maxValue = new Date();
            }
            else {
                minValue = new Number();
                maxValue = new Number();  
            }        
            minValue = d3.min(dp, function (d) { return d[0] });
            maxValue = d3.max(dp, function (d) { return d[0] });

            this.controlChartViewModel.minX = minValue;
            this.controlChartViewModel.maxX = maxValue;
        }

        private PlotData() {
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            let data = viewModel.data;
            let points = viewModel.points;
            // Line      
            var xScale = this.xScale;
            var yScale = this.yScale;
            var d3line2 = d3.svg.line()
                .x(function (d) { return xScale(d[0]) })
                .y(function (d) { return yScale(d[1]) });
            //add line
            this.svgGroupMain.append("svg:path").classed('trend_Line', true)
                .attr("d", d3line2(points))
                .style("stroke-width", '1.5px')
                .style({ "stroke": data.LineColor, "stroke-dasharray": (data.LineStyle) })
                .style("fill", 'none');
            //add dots
            var dots = this.svgGroupMain.attr("id", "groupOfCircles").selectAll("dot")
                .data(points)
                .enter().append("circle")
                .style("fill", data.DataColor)
                .attr("r", viewModel.marker.MarkerSize)
                //.attr("r", 2.5)
                .attr("cx", function (d) { return xScale(d[0]); })
                .attr("cy", function (d) { return yScale(d[1]); });          

            this.dots = dots;

            //add tooltip
            var xFormat;
            if (viewModel.isDateRange)
                xFormat = d3.time.format(viewModel.xAxis.AxisLabelFormat);
            else
                xFormat = d3.format(viewModel.xAxis.AxisLabelFormat);
            var yFormat = d3.format(viewModel.yAxis.AxisLabelFormat);
            this.tooltipServiceWrapper.addTooltip(dots,
                (tooltipEvent: TooltipEventArgs<number>) => ControlChart.getTooltipData(tooltipEvent.data, data.DataColor, xFormat, yFormat),
                (tooltipEvent: TooltipEventArgs<number>) => null);
        }

        private static getTooltipData(value: any, datacolor: string, xFormat: any, yFormat: any): VisualTooltipDataItem[] {
            return [{
                displayName: xFormat(value[0]).toString(),
                value: yFormat(value[1]).toString(),
                color: datacolor
            }];
        }

        private GetSubgroups() {
            let subgroup: Subgroup = {
                uCL: 0,
                lCL: 0,
                mean: 0,
                sum: 0,
                startX: null,
                endX: null,
                stage: '',
                count: 0,
                stageDividerX: null,
                firstId: 0,
                lastId: 0,
                mRError: false
            };
            this.subgroups = [];
            let stages: Subgroup[] = [];
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            let points = viewModel.points;

            var dataView = this.dataView;
            if (!dataView
                || !dataView
                || !dataView.categorical
                || !dataView.categorical.values)
                return [];

            var subgroupValue: string;
            var categorical;
            var hasSubgroups: boolean;
            if (dataView.categorical.values.length == 1 || !dataView.categorical.values[1]) {
                subgroupValue = '';
                hasSubgroups = false;
            }
            else {
                categorical = dataView.categorical;
                subgroupValue = <string>categorical.values[1].values[0];
                hasSubgroups = true;
            }
            
            subgroup.startX = points[0][0];
            var currentSubgroup: string = subgroupValue;
            subgroup.stage = subgroupValue;
            for (let i = 0; i < points.length; i++) {
                var obj = points[i];

                if (hasSubgroups)
                    subgroupValue = <string>categorical.values[1].values[i];

                if (currentSubgroup == subgroupValue || !hasSubgroups) {
                    subgroup.sum = subgroup.sum + obj[1];
                    subgroup.count++;
                }
                else {
                    if (subgroup.count > 0)
                        subgroup.mean = subgroup.sum / subgroup.count;
                    if (i > 0)
                        subgroup.endX = points[i - 1][0];
                    else
                        subgroup.endX = obj[0];

                    var nextStartDate: any = <any>(obj[0]);
                    var endDate: any = (points[i - 1][0]);
                    subgroup.stageDividerX = ((nextStartDate.valueOf() + endDate.valueOf()) / 2);

                    stages.push({
                        startX: subgroup.startX,
                        endX: subgroup.endX,
                        count: subgroup.count,
                        mean: subgroup.mean,
                        sum: subgroup.sum,
                        uCL: 0,
                        lCL: 0,
                        stage: subgroup.stage,
                        stageDividerX: subgroup.stageDividerX,
                        firstId: subgroup.firstId,
                        lastId: i - 1,
                        mRError: false
                    });
                    subgroup.mean = obj[1];
                    subgroup.count = 1;
                    subgroup.sum = obj[1];
                    subgroup.startX = obj[0];
                    subgroup.firstId = i;
                    subgroup.stage = subgroupValue;
                }
                //set stage to last stage value
                if (hasSubgroups)
                    currentSubgroup = <string>categorical.values[1].values[i];

                //last point
                if (i == (points.length - 1)) {
                    if (subgroup.count > 0)
                        subgroup.mean = subgroup.sum / subgroup.count;
                    subgroup.endX = obj[0];
                    subgroup.stage = subgroupValue;
                    stages.push({
                        startX: subgroup.startX,
                        endX: subgroup.endX,
                        count: subgroup.count,
                        mean: subgroup.mean,
                        sum: subgroup.sum,
                        uCL: 0,
                        lCL: 0,
                        stage: subgroup.stage,
                        stageDividerX: subgroup.endX,
                        firstId: subgroup.firstId,
                        lastId: i,
                        mRError: false
                    });
                }
            }
            this.subgroups = stages;
        }

        private PlotMean() {
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            let mLine = viewModel.meanLine;
            var xScale = this.xScale;
            var yScale = this.yScale;
            var meanLine = this.meanLine;

            var mean = this.svgGroupMain.selectAll("meanLine")
                .data(meanLine)
                .enter().append("polyline")
                .attr("points", function (d) { return xScale(d['x1']).toString() + "," + yScale(d['y2']).toString() + "," + xScale(d['x2']).toString() + "," + yScale(d['y2']).toString() })
                .style({"stroke": mLine.lineColor, "stroke-width": 1.5, "stroke-dasharray": (mLine.lineStyle) });

            // x&#x0304; - symbol for x-bar
            var xbar = 0x0304;
            var yformatValue = d3.format(viewModel.yAxis.AxisLabelFormat);

            var stages = this.subgroups;
            var plot = this.plot;
            var meanText = this.svgGroupMain.selectAll("meanText")
                .data(meanLine)
                .enter().append("text")
                .attr("x", function (d) { if (stages.length == 1) return plot.width; else { return xScale(d['x1']).toString() } })
                .attr("y", function (d) { return yScale(d['y2']).toString() })
                .attr("dx", ".15em")
                .attr("dy", function (d) { if (stages.length == 1) return ".30em"; else return "-.25em" })
                .attr("text-anchor", "start")
                .text(function (d) { return 'x' + String.fromCharCode(xbar) + ' = ' + yformatValue(d['y2']).toString() })
                .style("font-size", mLine.textSize + 'px')
                .style("font-family", mLine.textFont)
                .style("fill", mLine.textColor);

            //add tooltip
            var xFormat;
            if (viewModel.isDateRange)
                xFormat = d3.time.format(viewModel.xAxis.AxisLabelFormat);
            else
                xFormat = d3.format(viewModel.xAxis.AxisLabelFormat);

            this.tooltipServiceWrapper.addTooltip(mean,
                (tooltipEvent: TooltipEventArgs<number>) => ControlChart.getTooltipMeanData(tooltipEvent.data, mLine.textColor, 'Mean', yformatValue),
                (tooltipEvent: TooltipEventArgs<number>) => null);
            this.tooltipServiceWrapper.addTooltip(meanText,
                (tooltipEvent: TooltipEventArgs<number>) => ControlChart.getTooltipMeanData(tooltipEvent.data, mLine.textColor, 'Mean', yformatValue),
                (tooltipEvent: TooltipEventArgs<number>) => null);

        }

        private static getTooltipMeanData(value: any, datacolor: string, label: any, yFormat: any): VisualTooltipDataItem[] {
            return [{
                displayName: label,
                value: yFormat(value['y2']).toString(),
                color: datacolor
            }];
        }


        private CalcStats() {
            let stages = this.subgroups;
            var meanLine = [];
            var uclLines = [];
            var lclLines = [];
            var uCL: number;
            var lCL: number;
            var stageDividers = [];
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            let data = viewModel.points;
            var numsds = viewModel.standardDeviations;
            var mr: number = viewModel.movingRange;
            var d2: number = this.LookUpd2(viewModel.movingRange);
            var mrMin: number;
            var mrMax: number;
            var mrRangeSum: number;
            var rBar: number;
            var lCL: number;
            var uCL: number;

            for (let i = 0; i < stages.length; i++) {
                mrRangeSum = 0;
                if (mr > stages[i].count) {
                    //cannot process stage if moving range is greater than number of items in stage?
                    stages[i].mRError = true;
                    viewModel.mRError = true;
                }
                else {
                    for (let j = stages[i].firstId + mr - 1; j <= stages[i].lastId; j++) {
                        //for each data point in the stage
                        mrMin = data[j - 1][1];
                        mrMax = data[j - 1][1];
                        //move thru stage data points using mr window 
                        for (let k = j; k >= j - mr + 1; k--) {
                            if (mrMin > data[k][1])
                                mrMin = data[k][1];
                            if (mrMax < data[k][1])
                                mrMax = data[k][1];
                        }
                        mrRangeSum = mrRangeSum + Math.abs(mrMax - mrMin);
                    }
                    rBar = mrRangeSum / (stages[i].count - mr + 1);
                    uCL = stages[i].mean + numsds * rBar / d2;
                    lCL = stages[i].mean - numsds * rBar / d2;
                    stages[i].uCL = uCL;
                    stages[i].lCL = lCL;
                }
                if (i > 0) {
                    meanLine.push({ x1: stages[i - 1].stageDividerX, y1: stages[i - 1].mean, x2: stages[i].stageDividerX, y2: stages[i].mean });
                    stageDividers.push({ x1: stages[i].stageDividerX, prevDividerX: stages[i - 1].stageDividerX, stageName: stages[i].stage });
                    if (!stages[i].mRError) {
                        uclLines.push({ x1: stages[i - 1].stageDividerX, y1: uCL, x2: stages[i].stageDividerX, y2: uCL });
                        lclLines.push({ x1: stages[i - 1].stageDividerX, y1: lCL, x2: stages[i].stageDividerX, y2: lCL });
                    }
                }
                else {
                    //show left divider
                    stageDividers.push({ x1: stages[i].startX, prevDividerX: 0, stageName: stages[i].stage });

                    meanLine.push({ x1: stages[i].startX, y1: stages[i].mean, x2: stages[i].stageDividerX, y2: stages[i].mean });
                    stageDividers.push({ x1: stages[i].stageDividerX, prevDividerX: stages[i].startX, stageName: stages[i].stage });
                    if (!stages[i].mRError) {
                        uclLines.push({ x1: stages[i].startX, y1: uCL, x2: stages[i].stageDividerX, y2: uCL });
                        lclLines.push({ x1: stages[i].startX, y1: lCL, x2: stages[i].stageDividerX, y2: lCL });
                    }
                }
            }
            this.subgroups = stages;
            this.meanLine = meanLine;
            this.lclLines = lclLines;
            this.uclLines = uclLines;
            this.subgroupDividers = stageDividers;
        }

        private DrawMRWarning() {
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            var plot = this.plot;
            var yAxisOffset = 54;
            if (viewModel.mRError && viewModel.showmRWarning) {
                this.svgGroupMain.append("text")
                    .attr("y", yAxisOffset)
                    .attr("x", (plot.width / 2))
                    .style("text-anchor", "middle")
                    .style("font-size", '12px')
                    .style("fill", 'red')
                    .text('Selected Moving Range is greater than the number of data points in a Subgroup');
            }
        }

        private DrawSubgroupDividers() {
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            let stageDiv = viewModel.subgroupDividerLine;
            var plot = this.plot;
            var stageDividers = this.subgroupDividers;
            var xScale = this.xScale;

            if(stageDividers.length > 1){ //only show where there are more than one stage
                this.svgGroupMain.selectAll("divider")
                    .data(stageDividers)
                    .enter().append("polyline")
                    .attr("points", function (d) {if(xScale(d['x1']) < plot.width) return xScale(d['x1']).toString() + "," + "0," + xScale(d['x1']).toString() + "," + plot.height.toString() })
                    .style({ "stroke": stageDiv.lineColor, "stroke-width": 1.5, "stroke-dasharray": (stageDiv.lineStyle) });
                var dividerText = this.svgGroupMain.selectAll("dividerText")
                    .data(stageDividers)
                    .enter().append("text")
                    .attr("x", function (d) { return xScale(d['prevDividerX']).toString()})
                    .attr("y", "0")
                    .attr("dx", ".35em")
                    .attr("dy", "1em")
                    .attr("text-anchor", "start")
                    .text(function (d) {if(xScale(d['prevDividerX']) >0) return d['stageName'].toString()  })
                    .style("font-size", stageDiv.textSize + 'px')
                    .style("font-family", stageDiv.textFont)
                    .style("fill", stageDiv.textColor);
            }

            this.tooltipServiceWrapper.addTooltip(dividerText,
                (tooltipEvent: TooltipEventArgs<number>) => ControlChart.getTooltipDividerText(tooltipEvent.data, stageDiv.textColor),
                (tooltipEvent: TooltipEventArgs<number>) => null);

        }

        private static getTooltipDividerText(value: any, datacolor: string): VisualTooltipDataItem[] {
            return [{
                displayName: 'Subgroup',
                value: value['stageName'].toString(),
                color: datacolor
            }];
        }

        private PlotControlLimits() {
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            let limitLine = viewModel.limitLine;
            var lclLines = this.lclLines;
            var uclLines = this.uclLines;
            var xScale = this.xScale;
            var yScale = this.yScale;
            var uclLine = this.svgGroupMain.selectAll("uclLine")
                .data(uclLines)
                .enter().append("polyline")
                .attr("points", function (d) { return xScale(d['x1']).toString() + "," + yScale(d['y1']).toString() + "," + xScale(d['x2']).toString() + "," + yScale(d['y2']).toString() })
                .style({ "stroke": limitLine.lineColor, "stroke-width": 1.5, "stroke-dasharray": (limitLine.lineStyle) });
            var lclLine = this.svgGroupMain.selectAll("lclLine")
                .data(lclLines)
                .enter().append("polyline")
                .attr("points", function (d) { return xScale(d['x1']).toString() + "," + yScale(d['y1']).toString() + "," + xScale(d['x2']).toString() + "," + yScale(d['y2']).toString() })
                .style({ "stroke": limitLine.lineColor, "stroke-width": 1.5, "stroke-dasharray": (limitLine.lineStyle) });
            var yformatValue = d3.format(viewModel.yAxis.AxisLabelFormat);

            var stages = this.subgroups;
            var plot = this.plot;

            var uclText = this.svgGroupMain.selectAll("uclText")
                .data(uclLines)
                .enter().append("text")
                .attr("x", function (d) { if (stages.length == 1) return plot.width; else { return xScale(d['x1']).toString() } })
                .attr("y", function (d) { return yScale(d['y1']).toString() })
                .attr("dx", ".15em")
                .attr("dy", function (d) { if (stages.length == 1) return ".30em"; else return "-.25em" })
                .attr("text-anchor", "start")
                .text(function (d) { return 'UCL = ' + yformatValue(d['y1']).toString() })
                .style("font-size", limitLine.textSize + 'px')
                .style("font-family", limitLine.textFont)
                .style("fill", limitLine.textColor);

            var lclText = this.svgGroupMain.selectAll("lclText")
                .data(lclLines)
                .enter().append("text")
                .attr("x", function (d) { if (stages.length == 1) return plot.width; else { return xScale(d['x1']).toString() } })
                .attr("y", function (d) { return yScale(d['y1']).toString() })
                .attr("dx", ".15em")
                .attr("dy", function (d) { if (stages.length == 1) return ".30em"; else return ".95em" })
                .attr("text-anchor", "start")
                .text(function (d) { return 'LCL = ' + yformatValue(d['y1']).toString() })
                .style("font-size", limitLine.textSize + 'px')
                .style("font-family", limitLine.textFont)
                .style("fill", limitLine.textColor);

            //add tooltip
            var yFormat = d3.format(viewModel.yAxis.AxisLabelFormat);
            this.tooltipServiceWrapper.addTooltip(lclLine,
                (tooltipEvent: TooltipEventArgs<number>) => ControlChart.getTooltipLimitData(tooltipEvent.data, limitLine.textColor, 'LCL', yFormat),
                (tooltipEvent: TooltipEventArgs<number>) => null);

            this.tooltipServiceWrapper.addTooltip(uclLine,
                (tooltipEvent: TooltipEventArgs<number>) => ControlChart.getTooltipLimitData(tooltipEvent.data, limitLine.textColor, 'UCL', yFormat),
                (tooltipEvent: TooltipEventArgs<number>) => null);

            this.tooltipServiceWrapper.addTooltip(lclText,
                (tooltipEvent: TooltipEventArgs<number>) => ControlChart.getTooltipLimitData(tooltipEvent.data, limitLine.textColor, 'LCL', yFormat),
                (tooltipEvent: TooltipEventArgs<number>) => null);

            this.tooltipServiceWrapper.addTooltip(uclText,
                (tooltipEvent: TooltipEventArgs<number>) => ControlChart.getTooltipLimitData(tooltipEvent.data, limitLine.textColor, 'UCL', yFormat),
                (tooltipEvent: TooltipEventArgs<number>) => null);

        }

        private static getTooltipLimitData(value: any, datacolor: string, label: any, yFormat: any): VisualTooltipDataItem[] {
            return [{
                displayName: label,
                value: yFormat(value['y1']).toString(),
                color: datacolor
            }];
        }


        private ApplyRules() {
            //rule 1 - highlight over/below UCL/LCL
            var consecutiveUPoints = [];
            var consecutiveLPoints = [];
            var consecIncPoints = [];
            var consecDecPoints = [];
            var meanPoints = [];
            let stages = this.subgroups;
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            let data = viewModel.points;
            for (let i = 0; i < stages.length; i++) {
                for (let j = stages[i].firstId; j <= stages[i].lastId; j++) {
                    //Rule #1 - outside LCL and UCL
                    if (viewModel.runRule1)
                        if (stages[i].uCL < data[j][1] || stages[i].lCL > data[j][1])
                            meanPoints.push([j]);

                    //Rule #2 - over 5 incr or decr
                    if (viewModel.runRule2) {
                        if (j > 0) {
                            if (data[j].yValue > data[j - 1][1])
                                consecIncPoints.push([j])
                            else {
                                if (consecIncPoints.length > 5)
                                    this.DrawRulePoints(consecIncPoints);
                                consecIncPoints = [];
                                consecIncPoints.push([j]);
                            }
                            if (data[j].yValue < data[j - 1][1])
                                consecDecPoints.push([j])
                            else {
                                if (consecDecPoints.length > 5)
                                    this.DrawRulePoints(consecDecPoints);
                                consecDecPoints = [];
                                consecDecPoints.push([j]);
                            }
                        }
                        else {
                            consecIncPoints.push([j]);
                            consecDecPoints.push([j]);
                        }
                    }

                    //Rule #3 - Consecutive points in sequence greater than 8
                    if (viewModel.runRule3)
                        if (data[j].yValue > stages[i].mean) {
                            if (consecutiveLPoints.length > 8)
                                this.DrawRulePoints(consecutiveLPoints);
                            consecutiveLPoints = [];
                            consecutiveUPoints.push([j]);
                        }
                        else
                            if (data[j].yValue < stages[i].mean) {
                                if (consecutiveUPoints.length > 8)
                                    this.DrawRulePoints(consecutiveUPoints);
                                consecutiveUPoints = [];
                                consecutiveLPoints.push([j]);
                            }
                }
                if (meanPoints.length > 0) {
                    this.DrawRulePoints(meanPoints);
                    meanPoints = [];
                }
                if (consecutiveLPoints.length > 8)
                    this.DrawRulePoints(consecutiveLPoints);
                if (consecutiveUPoints.length > 8)
                    this.DrawRulePoints(consecutiveUPoints);
                consecutiveLPoints = [];
                consecutiveUPoints = [];

                if (consecIncPoints.length > 5)
                    this.DrawRulePoints(consecIncPoints);
                if (consecDecPoints.length > 5)
                    this.DrawRulePoints(consecDecPoints);
                consecDecPoints = [];
                consecIncPoints = [];
            }
        }

        private DrawRulePoints(points: any) {
            let dots = this.dots;
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            for (let i = 0; i < points.length; i++) {
                d3.select(dots[0][points[i]]).style("fill", viewModel.ruleColor);
            }
        }

        private LookUpd2(mr: number): number {
            var d2Array = [1,
                1.128,
                1.693,
                2.059,
                2.326,
                2.534,
                2.704,
                2.847,
                2.97,
                3.078,
                3.173,
                3.258,
                3.336,
                3.407,
                3.472,
                3.532,
                3.588,
                3.64,
                3.689,
                3.735,
                3.778,
                3.819,
                3.858,
                3.895,
                3.931,
                3.964,
                3.997,
                4.027,
                4.057,
                4.086,
                4.113,
                4.139,
                4.165,
                4.189,
                4.213,
                4.236,
                4.259,
                4.28,
                4.301,
                4.322,
                4.341,
                4.361,
                4.379,
                4.398,
                4.415,
                4.433,
                4.45,
                4.466,
                4.482,
                4.498];
            if (mr > 0)
                return d2Array[mr - 1];
            else
                return 1;
        }

       

        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] {
            var instances: VisualObjectInstance[] = [];
            var viewModel = this.controlChartViewModel;
            var objectName = options.objectName;
            switch (objectName) {
                case 'chart':
                    var config: VisualObjectInstance = {
                        objectName: objectName,
                        selector: null,
                        properties: {
                            dataColor: viewModel.data.DataColor,
                            lineColor: viewModel.data.LineColor,
                            lineStyle: viewModel.data.LineStyle,
                            backgroundColor: viewModel.backgroundColor,
                            markerSize: viewModel.marker.MarkerSize,
                            gridlinesColor: viewModel.gridlines.lineColor,
                            gridlinesStyle: viewModel.gridlines.lineStyle,
                            showGridLines: viewModel.gridlines.show
                        },
                        validValues: {
                            markerSize: {
                                numberRange: {
                                    min: 1,
                                    max: 20
                                }
                            }
                        }
                    };
                    instances.push(config);
                    break;
                case 'subgroups':
                    var config: VisualObjectInstance = {
                        objectName: objectName,
                        selector: null,
                        properties: {
                            showDividers: viewModel.subgroupDividerLine.show, 
                            subgroupDividerColor: viewModel.subgroupDividerLine.lineColor,
                            subgroupDividerLineStyle: viewModel.subgroupDividerLine.lineStyle,
                            subgroupLabelColor: viewModel.subgroupDividerLine.textColor,
                            subgroupDividerLabelSize: viewModel.subgroupDividerLine.textSize,
                            subgroupDividerfontFamily: viewModel.subgroupDividerLine.textFont,
                            showLimits: viewModel.limitLine.show,
                            limitLineColor: viewModel.limitLine.lineColor,
                            limitLineStyle: viewModel.limitLine.lineStyle,  
                            limitLabelColor: viewModel.limitLine.textColor,
                            limitLabelSize: viewModel.limitLine.textSize,
                            limitLabelfontFamily: viewModel.limitLine.textFont                                                     
                        },
                        validValues: {
                            subgroupDividerLabelSize: {
                                numberRange: {
                                    min: 4,
                                    max: 30
                                }
                            },
                            limitLabelSize: {
                                numberRange: {
                                    min: 4,
                                    max: 30
                                }
                            }
                        }
                    };
                    instances.push(config);
                    break;
                case 'statistics':
                    var config: VisualObjectInstance = {
                        objectName: objectName,
                        selector: null,
                        properties: {                            
                            showMean: viewModel.meanLine.show,
                            meanLineColor: viewModel.meanLine.lineColor,
                            meanLineStyle: viewModel.meanLine.lineStyle,
                            meanLabelColor: viewModel.meanLine.textColor,
                            meanLabelSize: viewModel.meanLine.textSize,
                            meanLabelfontFamily: viewModel.meanLine.textFont,                            
                            standardDeviations: viewModel.standardDeviations,
                            movingRange: viewModel.movingRange,
                            showmRWarning: viewModel.showmRWarning
                        },
                        validValues: {                            
                            meanLabelSize: {
                                numberRange: {
                                    min: 4,
                                    max: 30
                                }
                            },
                            movingRange: {
                                numberRange: {
                                    min: 2,
                                    max: 50
                                }
                            }
                        }
                    };
                    instances.push(config);
                    break;
                case 'xAxis':                 
                   
                    var dateformat: string = "%d-%b-%y";
                    var numericformat: string = ".3s";
                    if(viewModel.isDateRange){
                        dateformat = viewModel.xAxisFormat;                        
                           // numericformat = null;
                    }
                    else{
                        numericformat = viewModel.xAxisFormat;                        
                            //dateformat = null;
                    }

                    var config: VisualObjectInstance = {
                        objectName: objectName,
                        selector: null,
                        properties: {
                            xAxisTitle: viewModel.xAxis.AxisTitle,
                            xAxisTitleColor: viewModel.xAxis.TitleColor,
                            xAxisTitleSize: viewModel.xAxis.TitleSize,
                            xAxisTitlefontFamily: viewModel.xAxis.TitleFont,
                            xAxisLabelColor: viewModel.xAxis.AxisLabelColor,
                            xAxisLabelSize: viewModel.xAxis.AxisLabelSize,
                            xAxisLabelfontFamily: viewModel.xAxis.AxisLabelFont,                                
                            xAxisLabelFormat: numericformat,
                            xAxisDateFormat: dateformat,
                            xAxisLabelRotation: viewModel.xAxis.Rotation
                        },
                        validValues: {
                            xAxisTitleSize: {
                                numberRange: {
                                    min: 4,
                                    max: 30
                                }
                            },
                            xAxisLabelSize: {
                                numberRange: {
                                    min: 4,
                                    max: 30
                                }
                            },
                            xAxisLabelRotation: {
                                numberRange:  {
                                    min: 0,
                                    max: 360
                                }
                            }
                        }
                    };
                   
                    instances.push(config);
                    break;
                case 'yAxis':
                    var config: VisualObjectInstance = {
                        objectName: objectName,
                        selector: null,
                        properties: {
                            yAxisTitle: viewModel.yAxis.AxisTitle,
                            yAxisTitleColor: viewModel.yAxis.TitleColor,
                            yAxisTitleSize: viewModel.yAxis.TitleSize,
                            yAxisTitlefontFamily: viewModel.yAxis.TitleFont,
                            yAxisLabelColor: viewModel.yAxis.AxisLabelColor,
                            yAxisLabelSize: viewModel.yAxis.AxisLabelSize,                           
                            yAxisLabelfontFamily: viewModel.yAxis.AxisLabelFont,
                            yAxisLabelFormat: viewModel.yAxis.AxisLabelFormat
                        },
                        validValues: {
                            yAxisTitleSize: {
                                numberRange: {
                                    min: 4,
                                    max: 30
                                }
                            },
                            yAxisLabelSize: {
                                numberRange: {
                                    min: 4,
                                    max: 30
                                }
                            }
                        }
                    };
                    instances.push(config);
                    break;
                case 'rules':
                    var config: VisualObjectInstance = {
                        objectName: objectName,
                        selector: null,
                        properties: {
                            runRule1: viewModel.runRule1,
                            runRule2: viewModel.runRule2,
                            runRule3: viewModel.runRule3,
                            ruleColor: viewModel.ruleColor
                        }
                    };
                    instances.push(config);
                    break;
            }
            return instances;
        }
    }
}