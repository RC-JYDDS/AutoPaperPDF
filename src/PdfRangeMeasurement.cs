using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.Content.Objects;

namespace AutoPaper
{
    // Four differently coloured outer cell borders, printed only on a temporary
    // worksheet, measure the actual cell range in PDF coordinates. This keeps
    // intentional empty rows/columns and does not infer the range from ink.
    public static class PdfRangeMeasurement
    {
        sealed class State
        {
            public double A=1, B, C, D=1, E, F, LineWidth=1;
            public int Fill=-1, Stroke=-1;
            public State Copy() { return (State)MemberwiseClone(); }
            public XPoint Point(double x,double y) { return new XPoint(A*x+C*y+E,B*x+D*y+F); }
            public void Transform(double[] m)
            {
                double a=A*m[0]+C*m[1], b=B*m[0]+D*m[1], c=A*m[2]+C*m[3], d=B*m[2]+D*m[3];
                double e=A*m[4]+C*m[5]+E, f=B*m[4]+D*m[5]+F;
                A=a; B=b; C=c; D=d; E=e; F=f;
            }
        }
        sealed class Piece { public double Position, Start, End; }
        static double Number(CObject obj)
        {
            var i=obj as CInteger; if(i!=null)return i.Value;
            var r=obj as CReal; if(r!=null)return r.Value;
            throw new InvalidDataException("PDF 边界坐标无法识别。");
        }
        static int Color(double[] v, int[] colors)
        {
            if(v.Length==4) v=new[]{1-Math.Min(1,v[0]+v[3]),1-Math.Min(1,v[1]+v[3]),1-Math.Min(1,v[2]+v[3])};
            if(v.Length!=3) return -1;
            for(int i=0;i<4;i++)
                if(Math.Abs(v[0]*255-(colors[i]&255))<1.1 && Math.Abs(v[1]*255-((colors[i]>>8)&255))<1.1 &&
                   Math.Abs(v[2]*255-((colors[i]>>16)&255))<1.1) return i;
            return -1;
        }
        static int ColorOperands(COperator op,int[] colors)
        {
            if(op.Operands.Any(v=>!(v is CInteger) && !(v is CReal)))return -1;
            return Color(op.Operands.Select(Number).ToArray(),colors);
        }
        public static int[] CreateColors()
        {
            var random=new Random(Guid.NewGuid().GetHashCode()); var result=new List<int>();
            while(result.Count<4)
            {
                int r=random.Next(32,225),g=random.Next(32,225),b=random.Next(32,225);
                if(result.Any(c=>Math.Abs((c&255)-r)+Math.Abs(((c>>8)&255)-g)+Math.Abs(((c>>16)&255)-b)<100))continue;
                result.Add(r+(g<<8)+(b<<16));
            }
            return result.ToArray();
        }
        static XPoint Display(PdfPage page,XPoint p)
        {
            var box=page.MediaBox; double u=PdfPageFitter.PageUnit(page); int r=PdfPageFitter.PageRotation(page);
            switch(r)
            {
                case 90:return new XPoint((p.Y-box.Y1)*u,(p.X-box.X1)*u);
                case 180:return new XPoint((box.X2-p.X)*u,(p.Y-box.Y1)*u);
                case 270:return new XPoint((box.Y2-p.Y)*u,(box.X2-p.X)*u);
                default:return new XPoint((p.X-box.X1)*u,(box.Y2-p.Y)*u);
            }
        }
        static void Paint(PdfPage page,List<XPoint> path,int color,List<Piece>[] edges)
        {
            if(color<0 || path.Count<2)return;
            var points=path.Select(p=>Display(page,p)).ToArray();
            double x0=points.Min(p=>p.X),x1=points.Max(p=>p.X),y0=points.Min(p=>p.Y),y1=points.Max(p=>p.Y);
            bool vertical=color==0 || color==2;
            double thickness=vertical?x1-x0:y1-y0, length=vertical?y1-y0:x1-x0;
            if(thickness>3 || length<0.05)throw new InvalidDataException("边界测量标记不是直线，未裁切 PDF。");
            edges[color].Add(new Piece{Position=vertical?(x0+x1)/2:(y0+y1)/2,Start=vertical?y0:x0,End=vertical?y1:x1});
        }
        // An Office outer border may straddle the measured cell boundary.
        // Include its original ink only where it would cross the requested page.
        sealed class BorderInk
        {
            readonly XRect frame;
            public double Left, Top, Right, Bottom;
            public BorderInk(XRect frame, PageMargins margins)
            {
                this.frame = frame;
                Left = frame.Left - PageMargins.ToPoints(margins.LeftMm);
                Top = frame.Top - PageMargins.ToPoints(margins.TopMm);
                Right = frame.Right + PageMargins.ToPoints(margins.RightMm);
                Bottom = frame.Bottom + PageMargins.ToPoints(margins.BottomMm);
            }
            public void Include(PdfPage page, List<XPoint> path, double strokeRadius)
            {
                if (path.Count < 2) return;
                var pts = path.Select(p => Display(page, p)).ToArray();
                double x0 = pts.Min(p => p.X), x1 = pts.Max(p => p.X);
                double y0 = pts.Min(p => p.Y), y1 = pts.Max(p => p.Y);
                if (x0-strokeRadius >= Left && x1+strokeRadius <= Right &&
                    y0-strokeRadius >= Top && y1+strokeRadius <= Bottom) return;
                // Office prints borders as stroked lines/rectangles or thin filled strips.
                if (pts.Any(p => (p.X != x0 && p.X != x1) || (p.Y != y0 && p.Y != y1))) return;
                double reach = strokeRadius;
                if (strokeRadius == 0)
                {
                    reach = Math.Min(x1-x0, y1-y0);
                    if (reach > 6) return; // exclude cell and paper background fills
                }
                double near = reach + 0.6; // positioning tolerance used by the marker reader
                if (y1 >= frame.Top && y0 <= frame.Bottom)
                {
                    if (Math.Abs(x0-frame.Left) <= near) Left = Math.Min(Left, x0-strokeRadius);
                    if (Math.Abs(x1-frame.Right) <= near) Right = Math.Max(Right, x1+strokeRadius);
                }
                if (x1 >= frame.Left && x0 <= frame.Right)
                {
                    if (Math.Abs(y0-frame.Top) <= near) Top = Math.Min(Top, y0-strokeRadius);
                    if (Math.Abs(y1-frame.Bottom) <= near) Bottom = Math.Max(Bottom, y1+strokeRadius);
                }
            }
        }
        static void Scan(PdfPage page,CSequence sequence,PdfDictionary resources,State state,int[] colors,List<Piece>[] edges,int depth,BorderInk ink,List<XPoint> anchors)
        {
            if(depth>12)throw new InvalidDataException("PDF 图形嵌套过深，未裁切 PDF。");
            var stack=new Stack<State>(); var path=new List<XPoint>();
            foreach(CObject item in sequence)
            {
                var op=item as COperator; if(op==null)continue;
                string code=op.OpCode.Name;
                if(code=="q"){stack.Push(state.Copy());continue;}
                if(code=="Q"){if(stack.Count==0)throw new InvalidDataException("PDF 图形状态不完整。");state=stack.Pop();continue;}
                if(code=="cm"){state.Transform(op.Operands.Select(Number).ToArray());continue;}
                if(code=="w"){state.LineWidth=Number(op.Operands[0]);continue;}
                if(code=="Tm" && anchors!=null)
                    anchors.Add(Display(page,state.Point(Number(op.Operands[4]),Number(op.Operands[5]))));
                if(code=="rg" || code=="k" || code=="sc" || code=="scn")
                { state.Fill=ColorOperands(op,colors);continue; }
                if(code=="RG" || code=="K" || code=="SC" || code=="SCN")
                { state.Stroke=ColorOperands(op,colors);continue; }
                if(code=="g"){state.Fill=-1;continue;} if(code=="G"){state.Stroke=-1;continue;}
                if(code=="m" || code=="l") {path.Add(state.Point(Number(op.Operands[0]),Number(op.Operands[1])));continue;}
                if(code=="re")
                {
                    double[] v=op.Operands.Select(Number).ToArray();
                    path.Add(state.Point(v[0],v[1]));path.Add(state.Point(v[0]+v[2],v[1]+v[3]));continue;
                }
                if(code=="c" || code=="v" || code=="y")
                { for(int i=0;i<op.Operands.Count;i+=2)path.Add(state.Point(Number(op.Operands[i]),Number(op.Operands[i+1])));continue; }
                bool both = code=="B" || code=="B*" || code=="b" || code=="b*";
                bool fill = both || code=="f" || code=="f*" || code=="F";
                bool stroke = both || code=="S" || code=="s";
                if (fill) Paint(page,path,state.Fill,edges);
                if (stroke) Paint(page,path,state.Stroke,edges);
                if (ink != null && (fill || stroke))
                {
                    double scale = Math.Sqrt(Math.Max(state.A*state.A+state.C*state.C, state.B*state.B+state.D*state.D));
                    ink.Include(page, path, stroke ? state.LineWidth*scale*PdfPageFitter.PageUnit(page)/2 : 0);
                }
                if (fill || stroke || code=="n") path.Clear();
                if(code=="Do" && resources!=null)
                {
                    var xo=resources.Elements.GetDictionary("/XObject"); var name=op.Operands[0] as CName;
                    var form=xo==null || name==null?null:xo.Elements.GetDictionary(name.Name);
                    if(form==null)continue;
                    if(form.Elements.GetName("/Subtype")=="/Image" && anchors!=null)
                    {
                        anchors.Add(Display(page,state.Point(0,0)));
                        anchors.Add(Display(page,state.Point(1,1)));
                    }
                    if(form.Elements.GetName("/Subtype")!="/Form")continue;
                    var nested=state.Copy();var matrix=form.Elements.GetArray("/Matrix");
                    if(matrix!=null)nested.Transform(Enumerable.Range(0,6).Select(i=>matrix.Elements.GetReal(i)).ToArray());
                    Scan(page,ContentReader.ReadContent(form.Stream.UnfilteredValue),form.Elements.GetDictionary("/Resources")??resources,nested,colors,edges,depth+1,ink,anchors);
                }
            }
        }
        static List<Piece>[] Pieces(PdfPage page,int[] colors,BorderInk ink=null,List<XPoint> anchors=null)
        {
            if(colors==null || colors.Length!=4 || colors.Distinct().Count()!=4)throw new ArgumentException("需要四种不同的测量颜色。");
            var result=Enumerable.Range(0,4).Select(i=>new List<Piece>()).ToArray();
            Scan(page,ContentReader.ReadContent(page),page.Elements.GetDictionary("/Resources"),new State(),colors,result,0,ink,anchors);return result;
        }
        public static void VerifyNoMarkers(PdfPage page,int[] colors)
        { CheckOriginal(Pieces(page,colors)); }
        static void CheckOriginal(List<Piece>[] pieces)
        {
            if (pieces.Any(p=>p.Count>0))
                throw new InvalidDataException("测量颜色与原表格冲突，未裁切 PDF；请重新导出。");
        }
        internal static XRect PageWithMargins(PdfPage page,int[] colors,XRect frame,PageMargins margins,List<XPoint> probeAnchors)
        {
            var ink = new BorderInk(frame, margins);
            var anchors = new List<XPoint>();
            CheckOriginal(Pieces(page,colors,ink,anchors));
            // Excel can paint unchanged cells in a different order on the copy.
            var original = anchors.OrderBy(p=>p.X).ThenBy(p=>p.Y);
            var probe = probeAnchors.OrderBy(p=>p.X).ThenBy(p=>p.Y);
            if (anchors.Count != probeAnchors.Count || original.Zip(probe,(a,b) =>
                Math.Abs(a.X-b.X)>0.5 || Math.Abs(a.Y-b.Y)>0.5).Any(different=>different))
                throw new PaperLayoutMismatchException("原始 PDF 与测量 PDF 的文字或图片位置不一致，未裁切 PDF。");
            return new XRect(ink.Left,ink.Top,ink.Right-ink.Left,ink.Bottom-ink.Top);
        }
        public static XRect Read(PdfPage page,int[] colors)
        { return Read(page,colors,null); }
        internal static XRect Read(PdfPage page,int[] colors,List<XPoint> anchors)
        {
            var pieces=Pieces(page,colors,null,anchors); var positions=new double[4];
            for(int i=0;i<4;i++)
            {
                if(pieces[i].Count==0)
                {
                    if((i==2 || i==3) && pieces[0].Count>0 && pieces[1].Count>0)
                        throw new PaperTooSmallException("右侧或底部边界未完整导出，正在尝试更大的中间纸型。");
                    throw new InvalidDataException("表格软件未导出完整的四边测量标记，未裁切 PDF。");
                }
                var list=pieces[i];positions[i]=list.Sum(p=>p.Position*(p.End-p.Start))/list.Sum(p=>p.End-p.Start);
                if(list.Any(p=>Math.Abs(p.Position-positions[i])>0.6))throw new InvalidDataException("打印区域边界不一致，未裁切 PDF。");
            }
            if(positions[2]<=positions[0] || positions[3]<=positions[1])throw new InvalidDataException("打印区域边界顺序无效。");
            for(int i=0;i<4;i++)
            {
                bool vertical=i==0 || i==2;double start=vertical?positions[1]:positions[0],end=vertical?positions[3]:positions[2];
                double covered=start;foreach(var p in pieces[i].OrderBy(p=>p.Start))
                { if(p.Start>covered+1.5)throw new InvalidDataException("打印区域边界不连续，未裁切 PDF。");covered=Math.Max(covered,p.End); }
                if(covered<end-1.5 || pieces[i].Min(p=>p.Start)<start-1.5 || covered>end+1.5)throw new InvalidDataException("打印区域边界没有闭合，未裁切 PDF。");
            }
            return new XRect(positions[0],positions[1],positions[2]-positions[0],positions[3]-positions[1]);
        }
    }
}
