from matplotlib.figure import Figure

def plot_text(ax, text, coordinates, font_size):
    for i, char in enumerate(text):
        if i < len(coordinates):
            x, y = list(coordinates)[i]
            ax.text(x, y, char,
                   fontsize=font_size, color="black",
                   ha='center', va='center',
                   fontweight='bold', fontname='Arial')

def create_preview_figure(image, source_coords, client_data, coordinate_maps, paper_dims):
    fig = Figure(figsize=(paper_dims['width'], paper_dims['height']), dpi=344)
    ax = fig.add_subplot(111)
    
    # Plot image and source code points
    ax.imshow(image)
    ax.scatter(*zip(*source_coords), color="red", s=30, marker="o")
    
    # Calculate font size
    font_size = 26 * (paper_dims['height'] / 14.0)
    
    # Plot all names
    plot_text(ax, client_data['first_name'], coordinate_maps['FirstName'].values(), font_size)
    plot_text(ax, client_data['middle_name'], coordinate_maps['MiddleName'].values(), font_size)
    plot_text(ax, client_data['last_name'], coordinate_maps['LastName'].values(), font_size)
    
    ax.axis('off')
    return fig

def create_print_figure(image, source_coords, client_data, coordinate_maps, paper_dims):
    fig = Figure(figsize=(paper_dims['width'], paper_dims['height']), dpi=344)
    ax = fig.add_subplot(111)
    
    ax.imshow(image)
    ax.scatter(*zip(*source_coords), color="red", s=30, marker="o")
    
    font_size = 8 * (paper_dims['height'] / 14.0)
    
    plot_text(ax, client_data['first_name'], coordinate_maps['FirstName'].values(), font_size)
    plot_text(ax, client_data['middle_name'], coordinate_maps['MiddleName'].values(), font_size)
    plot_text(ax, client_data['last_name'], coordinate_maps['LastName'].values(), font_size)
    
    ax.axis('off')
    return fig