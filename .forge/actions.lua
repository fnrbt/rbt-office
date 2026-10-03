return {
    select = function(ctx)
        if string.match(ctx.ref_name, "^refs/tags/v%d+%.%d+%.%d+") then
            return { "publish-npm" }
        end
        return {}
    end,
    ["publish-npm"] = function(_)
        return {
            output = "package",
            base = ".forge/publish-npm",
            packages = {
                publish = {
                    { registry = "41021897", ecosystem = "npm", mode = "explicit", prefix = "@rbt/office" },
                },
            },
        }
    end,
}
