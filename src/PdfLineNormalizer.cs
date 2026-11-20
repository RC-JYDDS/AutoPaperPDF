using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.Content.Objects;
using PdfSharp.Pdf.IO;

namespace AutoPaper
{
    public static class PdfLineNormalizer
    {
        sealed class State
        {
            public double A=1, B, C, D=1, E, F;
            public string Fill="g:0";
            public State Copy() { return (State)MemberwiseClone(); }
            public void Transform(double[] m)
            {
                double a=A*m[0]+C*m[1], b=B*m[0]+D*m[1];
                double c=A*m[2]+C*m[3], d=B*m[2]+D*m[3];
                double e=A*m[4]+C*m[5]+E, f=B*m[4]+D*m[5]+F;
                A=a; B=b; C=c; D=d; E=e; F=f;
            }
        }
        sealed class Line
        {
            public COperator Op;
            public bool Horizontal;
            public string Fill;
            public double X,Y,W,H,Sx,Sy,Thickness,X0,X1,Y0,Y1;
        }
        static double Number(CObject obj)
        {
            var integer=obj as CInteger; if(integer!=null) return integer.Value;
            var real=obj as CReal; if(real!=null) return real.Value;
            throw new InvalidDataException("PDF 中存在无法识别的线条坐标。");
        }
        static double[] Numbers(COperator op)
        { return op.Operands.Select(Number).ToArray(); }
        static List<Line> Candidates(CSequence sequence)
        {
            var state=new State(); var stack=new Stack<State>(); var result=new List<Line>();
            for(int i=0;i<sequence.Count;i++)
            {
                var op=sequence[i] as COperator; if(op==null) continue;
                string code=op.OpCode.Name;
                if(code=="q") { stack.Push(state.Copy()); continue; }
                if(code=="Q") { if(stack.Count>0) state=stack.Pop(); continue; }
                if(code=="cm") { state.Transform(Numbers(op)); continue; }
                if(code=="rg" || code=="g" || code=="k" || code=="sc" || code=="scn")
                { state.Fill=code+":"+op.Operands.ToString(); continue; }
                if(code!="re" || op.Operands.Count!=4 || i+1>=sequence.Count) continue;
                var paint=sequence[i+1] as COperator;
                if(paint==null || (paint.OpCode.Name!="f" && paint.OpCode.Name!="f*")) continue;
                // Only axis-aligned filled rectangles. Do not alter images,
                // text, stroked/dashed paths, clipping paths, or rotated art.
                if(Math.Abs(state.B)>0.000001 || Math.Abs(state.C)>0.000001 ||
                   Math.Abs(state.A)<0.000001 || Math.Abs(state.D)<0.000001) continue;
                double[] v=Numbers(op); if(v[2]<=0 || v[3]<=0) continue;
                double sx=Math.Abs(state.A),sy=Math.Abs(state.D);
                double width=v[2]*sx,height=v[3]*sy,t=Math.Min(width,height);
                if(t<0.90 || t>1.02 || Math.Max(width,height)<10) continue;
                double x0=v[0]*state.A+state.E,x1=(v[0]+v[2])*state.A+state.E;
                double y0=v[1]*state.D+state.F,y1=(v[1]+v[3])*state.D+state.F;
                result.Add(new Line{Op=op,Horizontal=width>height,Fill=state.Fill,
                    X=v[0],Y=v[1],W=v[2],H=v[3],Sx=sx,Sy=sy,Thickness=t,
                    X0=Math.Min(x0,x1),X1=Math.Max(x0,x1),Y0=Math.Min(y0,y1),Y1=Math.Max(y0,y1)});
            }
            return result;
        }
        static bool IsGridLine(Line line,List<Line> all)
        {
            var crossings=new HashSet<double>();
            foreach(var other in all)
            {
                if(other.Horizontal==line.Horizontal || other.Fill!=line.Fill) continue;
                Line h=line.Horizontal?line:other, v=line.Horizontal?other:line;
                double x=(v.X0+v.X1)/2,y=(h.Y0+h.Y1)/2;
                if(x>=h.X0-0.2 && x<=h.X1+0.2 && y>=v.Y0-0.2 && y<=v.Y1+0.2)
                    crossings.Add(Math.Round(line.Horizontal?x:y,2));
            }
            // A border must meet at least two perpendicular borders of the
            // same colour. Isolated rules and decorative bars are preserved.
            return crossings.Count>=2;
        }
        public static int NormalizeWpsPdf(string input,string output)
        {
            if(String.Equals(Path.GetFullPath(input),Path.GetFullPath(output),StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("PDF 修正须使用临时输出文件。");
            using(var document=PdfReader.Open(input,PdfDocumentOpenMode.Modify))
            {
                string creator=document.Info.Creator??"";
                if(creator.IndexOf("WPS",StringComparison.OrdinalIgnoreCase)<0) return 0;
                if(document.Internals.Catalog.Elements.ContainsKey("/Perms"))
                    throw new InvalidDataException("PDF 带有权限或签名信息，未处理线条。");
                int changed=0;
                foreach(PdfPage page in document.Pages)
                {
                    var content=ContentReader.ReadContent(page);
                    var candidates=Candidates(content);
                    var targets=candidates.Where(l=>IsGridLine(l,candidates)).ToList();
                    if(targets.Count<4 || !targets.Any(l=>l.Horizontal) || !targets.Any(l=>!l.Horizontal)) continue;
                    foreach(var line in targets)
                    {
                        double dx=(line.Thickness-0.5)/(2*line.Sx),dy=(line.Thickness-0.5)/(2*line.Sy);
                        double[] values={line.X+dx,line.Y+dy,line.W-2*dx,line.H-2*dy};
                        for(int i=0;i<4;i++) line.Op.Operands[i]=new CReal{Value=values[i]};
                        changed++;
                    }
                    page.Contents.ReplaceContent(content);
                }
                if(changed==0) return 0;
                var dimensions=document.Pages.Cast<PdfPage>().Select(p=>new[]{p.Width.Point,p.Height.Point}).ToList();
                document.Save(output);
                using(var check=PdfReader.Open(output,PdfDocumentOpenMode.Import))
                {
                    if(check.PageCount!=dimensions.Count) throw new InvalidDataException("PDF 页数核对失败。");
                    for(int i=0;i<dimensions.Count;i++)
                    {
                        if(Math.Abs(check.Pages[i].Width.Point-dimensions[i][0])>0.001 ||
                           Math.Abs(check.Pages[i].Height.Point-dimensions[i][1])>0.001)
                            throw new InvalidDataException("PDF 页面尺寸核对失败。");
                    }
                }
                return changed;
            }
        }
    }
}
