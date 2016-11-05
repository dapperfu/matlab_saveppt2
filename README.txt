matlab_saveppt2
===============
% SAVEPPT2 saves plots to PowerPoint.
% SAVEPPT2(save_file,<additional parameters>)
%  Saves the current Matlab figure window or Simulink model window to a PowerPoint
%  file designated by save_file.  If save_file is omitted, the user is prompted to enter
%  one via UIPUTFILE.  If the path is omitted from filespec, the
%  PowerPoint file is created in the current Matlab working directory.
%
% powerpoint_object=SAVEPPT2(save_file,'close',false);
%   Return the PowerPoint presentation object if it isn't to be closed.
%
% Batch Processing:
% powerpoint_object=SAVEPPT2(save_file,'init')
%   Initializes a powerpoint object for use with batch processing.
% SAVEPPT2('ppt',powerpoint_object);
%   Saves the current Matlab figure to the opened PowerPoint Object
% SAVEPPT2(save_file,'ppt',powerpoint_object,'close');
%   Saves and closes the PowerPoint object.
%
% SAVEPPT2 also accepts numerous additional optional parameters, they can
% be called from the matlab command line or in function form. All options
% can be preceded with a '-', but this is not required. Short and Long
% named options are listed on the same line.
%
% % Annotation:
% 'n' 'notes'          - Add information to notes section.
% 'text' 'textbox'     - Add text box.
% 'comment' 'comments' - Add comment. (Only works if PowerPoint is visible)
%     - \t and \n are converted to tab and new line, respectively.
% 't' 'title' - Add a title or add a blank title so that one may be added later. Title is placed at the top of the presentation unless a padding is specified.
% If 'title' or 'textbox' is specified alone a blank placeholder will be added.
%
% % Figure Options
% 'f' 'fig' 'figure'     - Use the specified figure handle. Also accepts an array of figures. More than 4 figures is not recommended as it makes it difficult to see in the plot. Default: gcf
%                          If figure is 0, a blank page is added. If a title is specified then a title page is added.
% 'd' 'driver' 'drivers' - [meta, bitmap]. Send figure to clipboard Metafile or Bitmap format. See also print help.
% 'r' 'render'           - [painters,zbuffer,opengl]. Choose print render mode. See also print help.
% 'res' 'resolution'     - Dots-per-inch resolution. Default: 90 for Simulink, 150 for figures. See also print help.
%
% % Slide Layout
% 'st' 'stretch'      - Used only with scale, stretch the figure to fill all remaining space (taking into account padding and title). Default: on
% 's' 'sc' 'scale'    - Scale the figure to remaining space on the page while maintaining aspect ratio, takes into account padding and title spacing. Default: on
% 'h' 'halign'        - ['left','center','right']. Horizontally align figure. Default: center
% 'v' 'valign'        - ['top','center','bottom']. Vertically align the graph. Default: center
% 'p' 'pad' 'padding' - Place a padding around the figure that is used for alignment and scaling. Can be one number to be applied equally or an array in the format of [left right top bottom]. Useful when plotting to template files. Default: 0
% 'c' 'col' 'columns' - Number of columns to place multiple plots in. Default: 2
%
% % PowerPoint Control
% 'i' 'init' - Initialize PowerPoint presentation for use in batch mode. Returns a PowerPoint Presentation Object.
% 'close'    - Close PowerPoint presentation. Default: true
% 'save'     - Save PowerPoint Presentation. Useful for saves in batch mode.
% 'ppt'      - Call saveppt2 with specified PowerPoint Presentation object.
% 'visible'  - Make PowerPoint visible.
% 'template' - Use template file specified. Is only used if the save file does not already exist.
%
% For binary options use: 'yes','on' ,'true' ,true  to enable
%                         'no', 'off','false',false to disable
% Examples:
% % Simplest Call
% plot(rand(1,100),rand(1,100),'*');
% saveppt2
%
% % Add a title "Hello World"
% plot(rand(1,100),rand(1,100),'*');
% saveppt2('test.ppt','title','Hello World');
% saveppt2('test.ppt','t','Hello World');
% saveppt2 test.ppt -title 'Hello World'
% saveppt2 test.ppt -t 'Hello World'
%
% % Add a note
% plot(rand(1,100),rand(1,100),'*');
% saveppt2('test.ppt','notes','Lorem ipsum dolor sit amet, consectetur adipiscing elit.');
%
% % Add multiline note
% plot(rand(1,100),rand(1,100),'*');
% saveppt2('test.ppt','notes','Lifetime, relational competence.\n\tTactical, integrated solution');
% note=sprintf('Latest Quarter Sales: %%fM',rand(1,1)*100);
% saveppt2('test.ppt','notes',note);
%
% Add a TextBox
% plot(rand(1,100),rand(1,100),'*');
% saveppt2('test.ppt','texbox','Totally, groupwide mindset');
%
% % Add a comment (PowerPoint must be visible)
% plot(rand(1,100),rand(1,100),'*');
% saveppt2('test.ppt','visible','comment','Virtual, logic-based culture');
%
% % Scaling & Stretching the plot to fill the page.
% plot(rand(1,100),rand(1,100),'*');
% saveppt2('test.ppt','note','Scaling & Stretching On (Default)');
% saveppt2('test.ppt','stretch','false','Stretching Off');
% saveppt2('test.ppt','scale',off,'note','Scaling Off');
%
% % Stretching the plot to fill the page.
% plot(rand(1,100),rand(1,100),'*');
% saveppt2('test.ppt','Stretching On');
%
% saveppt2('test.ppt','stretch',false);
% saveppt2 test.ppt -stretch off
%
% % copy the plot as both a meta and bitmap.
% plot(rand(1,100),rand(1,100),'*');
% saveppt2('test.ppt','driver','meta','scale','stretch');
% saveppt2('test.ppt','driver','bitmap','scale','stretch');
%
% % scale the plot to fill the page, ignoring aspect ratio, with 150 pixels
% % of padding on all sides
% plot(rand(1,100),rand(1,100),'*');
% saveppt2('test.ppt','scale','stretch','Padding',150);
% saveppt2('test.ppt','scale',true,'stretch',true,'Padding',150);
%
% % scale the plot to fill the page, ignoring aspect ratio, with 150 pixels
% % of padding on left and right sides
% plot(rand(1,100),rand(1,100),'*');
% saveppt2('test.ppt','scale','stretch','Padding',[150 150 0 0]);
% saveppt2('test.ppt','scale',true,'stretch',true,'Padding',[150 150 0 0]);
%
% % scale the plot to fill the page, ignoring aspect ratio add blank title
% plot(rand(1,100),rand(1,100),'*');
% saveppt2('test.ppt','scale','stretch','title');
% saveppt2('test.ppt','scale',true,'stretch',true,'title',true);
%
% % Align the figure in the upper left corner
% plot(rand(1,100),rand(1,100),'*');
% saveppt2('test.ppt','halign','left','valign','top');
%
% % Align the figure in the upper left corner
% plot(rand(1,100),rand(1,100),'*');
% saveppt2('test.ppt','halign','right','valign','bottom');
%
% % Use the template 'Group Report.ppt'
% plot(rand(1,100),rand(1,100),'*');
% saveppt2('test.ppt','template','Group Report.ppt');
%
% % Plot 4 figures horizontally aligned left with 2 columns
% a=figure('Visible','off');plot(1:10);
% b=figure('Visible','off');plot([1:10].^2);
% c=figure('Visible','off');plot([1:10].^3);
% d=figure('Visible','off');plot([1:10].^4);
% saveppt2('test.ppt','figure',[a b c d],'columns',2,'title','Hello World!','halign','left')
%
% % Create blank title page.
% figure('test.ppt','figure',0,'title','New Section');
%
% % Create blank page.
% figure('test.ppt','figure',0);
%
% % Plot figures in batch mode. Faster than opening a new powerpoint object each time
% ppt=saveppt2('batch.ppt','init');
% for i=1:10
%     plot(rand(1,100),rand(1,100),'*');
%     saveppt2('ppt',ppt)
%     if mod(i,5)==0 % Save half way through incase of crash
%       saveppt2('ppt',ppt,'save')
%     end
% end
% saveppt2('batch.ppt','ppt',ppt,'close');
%
% More flexibility is built in, but it is impossible to show all possible
% calling combinations, you may check out the source or Test_SavePPT2.m
%
% See also print, saveppt, validateInput
