/// Sparse sheet-axis coordinates in pixels, shared by the spreadsheet model and writer.
module AxisGeometry

/// Prefix position of an index, with hidden rows collapsed to zero.
let prefix (defaultSize: float) (sizes: Map<int, float>) (hidden: int array) =
    let defaultSize = max 0.0 defaultSize
    let sizes = hidden |> Array.fold (fun map index -> Map.add index 0.0 map) sizes
    let keys = Array.zeroCreate<int> sizes.Count
    let sums = Array.zeroCreate<float> (sizes.Count + 1)
    let mutable i = 0
    sizes |> Map.iter (fun index size ->
        keys.[i] <- index
        sums.[i + 1] <- sums.[i] + max 0.0 size - defaultSize
        i <- i + 1)
    fun index ->
        let mutable lo, hi = 0, keys.Length
        while lo < hi do
            let mid = lo + (hi - lo) / 2
            if keys.[mid] < index then lo <- mid + 1 else hi <- mid
        float index * defaultSize + sums.[lo]

/// Find the cell and pixel offset holding a physical coordinate.
let locate (prefix: int -> float) limit position =
    let position = max 0.0 position
    let mutable lo, hi = 0, limit
    while lo + 1 < hi do
        let mid = lo + (hi - lo) / 2
        if prefix mid <= position then lo <- mid else hi <- mid
    lo, position - prefix lo
