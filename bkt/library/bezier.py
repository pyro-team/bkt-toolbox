# -*- coding: utf-8 -*-


def _punktAufGerade(t, P, Q):
    '''Berechne (Px,Py) + t*(Qx-Px,Qy-Py)'''
    return [P[0]+t*(Q[0]-P[0]), P[1]+t*(Q[1]-P[1])]

# _bezierCasteljau([[1,0],[1,0.5522],[0.5522,1],[0,1]], 0.5)
def _bezierCasteljau(bezier, t):
    '''Fuer eine Bezierkurve den Punkt fuer t\in[0,1] berechnen (nach Algorithmus von Casteljau)'''
    punkte = [bezier]
    for k in [1,2,3]:
        punkte.append([])
        for i in range(0,3-k+1):
            # punkte[k,i] = ...
            punkte[k].append(_punktAufGerade(t, punkte[k-1][i], punkte[k-1][i+1]))
    return punkte[3][0], punkte;

# Test:
#   vk = bezierKurveViertelkreis()
#   b1, b2 = _bezierKurveAnPunktTeilen(vk, 0.5)
#   x1 = _bezierCasteljau(vk,0.25)[0]
#   x2 = _bezierCasteljau(b1,0.5)[0]
#   x1[0]-x2[0] < 0.00001 and x1[1]-x2[1] < 0.00001
def _bezierKurveAnPunktTeilen(bezier, t):
    '''Bezierkurve an Kurvenpunkt fuer Parameter t\in[0,1] teilen.
       Liefert zwei neue Bezierkurven'''
    x, punkte = _bezierCasteljau(bezier, t);
    return [punkte[0][0], punkte[1][0], punkte[2][0], punkte[3][0]], [punkte[3][0], punkte[2][1], punkte[1][2], punkte[0][3]]

def _bezierKurveNFachTeilen(bezier, n):
    '''Bezierkurve bezier in n gleichlange Abschnitte teilen'''
    if (n <= 1):
        return [bezier]
    else:
        # beim ersten n-tel aufteilen
        b1, b2 = _bezierKurveAnPunktTeilen(bezier, 1.0/n);
        # rest (n-1)-mal aufteilen
        liste = [b1]
        liste.extend(_bezierKurveNFachTeilen(b2, n-1))
        return liste

def _bezierKurvenEinheitskreis():
    '''Einheitskreises als Liste von Bezierkurven'''
    return [[[1,0],[1,0.5522],[0.5522,1],[0,1]],
            [[0,1],[-0.5522,1],[-1,0.5522],[-1,0]],
            [[-1,0],[-1,-0.5522],[-0.5522,-1],[0,-1]],
            [[0,-1],[0.5522,-1],[1,-0.5522],[1,0]]];

# # Bezierdarstellung des ersten Viertels vom Einheitskreis
# def bezierKurveViertelkreis():
#     return [[1,0],[1,0.5522],[0.5522,1],[0,1]];

def _bezierKreisN(n):
    '''Einheitskreis in Bezierkurven, Quadranten n-fach geteilt'''
    kurven = map(lambda k : _bezierKurveNFachTeilen(k,n), _bezierKurvenEinheitskreis())
    return [item for sublist in kurven for item in sublist]

def _bezierKreisNR(n,r):
    '''Kreis mit Radius r um Nullpunkte in Bezierkurven, Quadranten n-fach geteilt'''
    kurven = _bezierKreisN(n)
    return [ [ [r*P[0], r*P[1]] for P in k] for k in kurven]

def bezierKreisNRM(n,r,M):
    '''Kreis mit Radius r um Punkt M=[x,y] in Bezierkurven, Quadranten n-fach geteilt'''
    kurven = _bezierKreisNR(n,r)
    return [ [ [P[0]+M[0], P[1]+M[1]] for P in k] for k in kurven]


def kreisSegmente(n,r,M):
    '''Kreis mit Radius r um Punkt M=[x,y], aufgeteilt in n Segmente, welche jeweils
     aus Bezierkurven zusammengesetzt sind. Liefert Liste dieser n Segmente.'''
    viertelKreise = bezierKreisNRM(1,r,M)
    
    viertelKreisRest = 0
    segmente = []
    
    # segmente nacheinander aufbauen
    for i in range(0,n):
        segmente.append([])
        # Teile Kreis in 4*n Teile
        # --> je Segment werden 4 Teile benoetigt, Viertelkreis entspricht n Teile
        # Kreisanteile, die noch fuer aktuelles Segment benoetigt werden
        segmentRest = 4;
        
        while segmentRest > 0:
            if viertelKreisRest == 0:
                # hole naechsten Viertelkreis
                restVomViertelKreis = viertelKreise.pop(0)
                viertelKreisRest = n # Anteile, die restVomViertelKreis representiert
                
            # nehme Teil vom Viertelkreis fuer aktuelles Segment
            # --> Anzahl Teile = min(viertelKreisRest, segmentRest)
            x = min(viertelKreisRest, segmentRest)
            # --> Parameter bezieht sich auf Laenge von Viertelkreis-Restkurve
            t = x*1./viertelKreisRest # \in[0,1]
            # restVomVierteilkreis wird bei t\in[0,1] aufgeteilt
            bezierFuerSegment, restVomViertelKreis = _bezierKurveAnPunktTeilen(restVomViertelKreis, t)
            # x Teile fuer aktuelles Segment
            segmente[i].append(bezierFuerSegment)
            segmentRest = segmentRest-x
            # Anzahl Teile im Rest vom Viertelkreis
            viertelKreisRest = viertelKreisRest-x
            
            # jetzt: Segment fertig (segmentRest=0) oder Viertelkreis zu Ende (viertelKreisRest=0)
            assert segmentRest==0 or viertelKreisRest==0
    return segmente



# TODO: in anderes Modul verschieben
def shapeAusKurvenSegmentListe(slide, kurvenSegmentListe):
    '''Erstelle neues Shape auf Slide aus geschachtelter Bezierkurvenliste: [[bezier1, bezier2],[bezier3, bezier4]]'''
    # erster Punkt
    P = kurvenSegmentListe[0][0][0];
    ffb = slide.Shapes.BuildFreeform(1, P[0], P[1])
    
    for segment in kurvenSegmentListe:
        for k in segment:
            # von den naechsten Beziers immer die naechsten Punkte angeben
            ffb.AddNodes(1, 3, k[1][0], k[1][1], k[2][0], k[2][1], k[3][0], k[3][1])
            # Parameter: SegmentType, EditingType, X1,Y1, X2,Y2, X3,Y3
            # SegmentType: 0=Line, 1=Curve
            # EditingType: 0=Auto, 1=Corner (keine Verbindungspunkte), 2=Smooth, 3=Symmetric  --> Zweck?
    return ffb.ConvertToShape()




